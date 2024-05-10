VERSION 5.00
Begin VB.Form frmCuent 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Tierras Perdidas AO"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCuent.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCuent.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   9015
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   9
      Left            =   8345
      MouseIcon       =   "frmCuent.frx":0CD6
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":19A0
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   47
      Top             =   4920
      Width           =   1250
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   9
         Left            =   0
         TabIndex        =   50
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   48
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   8
      Left            =   6840
      MouseIcon       =   "frmCuent.frx":1C3B
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":2905
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   42
      Top             =   4870
      Width           =   1250
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   46
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   8
         Left            =   0
         TabIndex        =   45
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   44
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   43
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   7
      Left            =   5280
      MouseIcon       =   "frmCuent.frx":2BA0
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":386A
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   37
      Top             =   4870
      Width           =   1250
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   7
         Left            =   0
         TabIndex        =   40
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   6
      Left            =   3720
      MouseIcon       =   "frmCuent.frx":3B05
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":47CF
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   32
      Top             =   4870
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   36
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   6
         Left            =   0
         TabIndex        =   34
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   5
      Left            =   2240
      MouseIcon       =   "frmCuent.frx":4A6A
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":5734
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   27
      Top             =   4870
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   29
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   4
      Left            =   8400
      MouseIcon       =   "frmCuent.frx":59CF
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":6699
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   22
      Top             =   3055
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   24
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   3
      Left            =   6840
      MouseIcon       =   "frmCuent.frx":6934
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":75FE
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   17
      Top             =   3100
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   1238
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1238
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1238
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   2
      Left            =   5280
      MouseIcon       =   "frmCuent.frx":7899
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":8563
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   12
      Top             =   3055
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   1
      Left            =   3720
      MouseIcon       =   "frmCuent.frx":87FE
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":94C8
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   7
      Top             =   3055
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   0
      Left            =   2240
      MouseIcon       =   "frmCuent.frx":9763
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":A42D
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   2
      Top             =   3055
      Width           =   1250
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   3360
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   6240
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   11160
      Top             =   120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   9000
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   480
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Accname"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   12840
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PJClick"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub CP_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)
If CP(Index).Visible = True Then
    Call Audio.PlayWave(SND_CLICK)

If Nombre(6).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Audio.StopWave
    Me.MousePointer = 11
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & "\Graficos\Icono.ico")
Dim cca As Integer
Me.Picture = LoadPicture(App.path & "\Graficos\Principal\Cuenta_Main.jpg")
cuentapj = 0
Dim i As Integer
'Label3.Caption = nombrecuent

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim asd As Integer
For asd = 0 To 6
Nombre(asd).ForeColor = vbWhite
    Label2(asd).ForeColor = vbWhite
Next asd
End Sub

Private Sub Image1_Click()
Call Audio.PlayWave(SND_CLICK)
'MsgBox "Desactivado momentaneamente."
'Exit Sub
If MsgBox("¿Estás seguro que deseas borrar a " & PJClickeado & "?", vbYesNo) = vbYes Then

    Call SendData("BORR" & PJClickeado & "," & nombrecuent)
    End If
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
If PJClickeado = "Nada" Or PJClickeado = "" Then
frmMensaje.Show
frmMensaje.MSG.Caption = "Seleccione un PJ."
Exit Sub
End If


 
SendData ("OOLOGI" & PJClickeado & "," & nombrecuent)

End Sub

Private Sub Image3_Click()
Call Audio.PlayWave(SND_CLICK)
frmMain.Socket1.Disconnect
frmMain.Socket1.Cleanup
Unload Me
frmConnect.Visible = True
End Sub

Private Sub Image4_Click()
Call Audio.PlayWave(SND_CLICK)
frmCambiarPass.Visible = True
Call SendData("PEDPRE" & nombrecuent)
End Sub

Private Sub Image5_Click()
Call Audio.PlayWave(SND_CLICK)
MsgBox "No habilitado momentaneamente."
End Sub

Private Sub Label2_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)
PJClickeado = Nombre(Index)
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Nombre(Index).ForeColor = vbYellow
    Label2(Index).ForeColor = vbYellow
End Sub

Private Sub nombre_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)
PJClickeado = Nombre(Index)
End Sub

Private Sub nombre_dblClick(Index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
SendData ("OOLOGI" & PJClickeado & "," & nombrecuent)

End Sub

Private Sub nombre_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Nombre(Index).ForeColor = vbYellow
    Label2(Index).ForeColor = vbYellow
End Sub

Private Sub PJ_Click(Index As Integer)
If CP(Index).Visible = True Then
    Call Audio.PlayWave(SND_CLICK)

If Nombre(6).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Audio.StopWave
    Me.MousePointer = 11
    Exit Sub
End If
PJClickeado = Nombre(Index)
End Sub

Private Sub PJ_dblClick(Index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
SendData ("OOLOGI" & PJClickeado & "," & nombrecuent)

End Sub

Private Sub PJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim asd As Integer
For asd = 0 To 6
Nombre(asd).ForeColor = vbWhite
    Label2(asd).ForeColor = vbWhite
Next asd
End Sub
