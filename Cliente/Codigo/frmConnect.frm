VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   210
      Left            =   5520
      TabIndex        =   2
      Top             =   4980
      Width           =   210
   End
   Begin VB.TextBox passwordtxt 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4575
      Width           =   3495
   End
   Begin VB.TextBox nametxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   0
      Top             =   3780
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recordar cuenta"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   4980
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   5160
      Top             =   6750
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   5160
      MousePointer    =   99  'Custom
      Top             =   5475
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   525
      Index           =   0
      Left            =   5160
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   2010
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 
Private Sub Check1_Click()
'Midraks
    If Check1.value = 0 Then
    Call WriteVar(App.Path & "\INIT\Recordar.dat", "Nombre", "Nombre", "")
    Call WriteVar(App.Path & "\INIT\Recordar.dat", "Password", "Password", "")
    Call WriteVar(App.Path & "\INIT\Recordar.dat", "Check", "Check", "0")
    nametxt.Text = ""
    passwordtxt.Text = ""
    MsgBox "Ha dejado de recordar su cuenta."
    End If
    '/Midraks
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 0, 0, 1, 0, 1
       
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
       
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        frmCargando.Refresh
        DeinitTileEngine
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
End If
End Sub
 
Private Sub Form_Load()

'Midraks
    If GetVar(App.Path & "\INIT\Recordar.dat", "Check", "Check") = 1 Then
    Check1.value = 1
    nametxt.Text = GetVar(App.Path & "\INIT\Recordar.dat", "Nombre", "Nombre")
    
    passwordtxt.Text = GetVar(App.Path & "\INIT\Recordar.dat", "Password", "Password")
    Else
    Check1.value = 0
    End If
    '/Midraks
Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Conectar.jpg")
 
End Sub
 
Private Sub Image1_Click(Index As Integer)
 
Call Audio.PlayWave("click.wav")
 
Select Case Index
    Case 0
             EstadoLogin = Dados
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
        frmMain.Socket1.HostName = MIIP
        frmMain.Socket1.RemotePort = MIPORT
        frmMain.Socket1.Connect
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        frmMain.Winsock1.Connect MIIP, MIPORT
#End If
        Me.MousePointer = 11
Case 1
If Check1.value = 1 Then
        Call WriteVar(App.Path & "\INIT\Recordar.dat", "Nombre", "Nombre", nametxt.Text)
        Call WriteVar(App.Path & "\INIT\Recordar.dat", "Password", "Password", passwordtxt.Text) 'Gracias a TonchitoZ por esto.
        Call WriteVar(App.Path & "\INIT\Recordar.dat", "Check", "Check", "1")
        End If
    #If UsarWrench = 1 Then
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    #Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
    #End If
            If frmConnect.MousePointer = 11 Then
                Exit Sub
            End If
           
           
            'update user info
            UserName = nametxt.Text
            Dim aux As String
            aux = passwordtxt.Text
    #If SeguridadAlkon Then
            UserPassword = md5.GetMD5String(aux)
            Call md5.MD5Reset
    #Else
            UserPassword = (aux)
    #End If
            If CheckUserData(False) = True Then
                'SendNewChar = False
                EstadoLogin = Normal
                Me.MousePointer = 11
                frmMain.Socket1.Disconnect
                frmConnect.MousePointer = vbNormal

                frmMain.Socket1.HostName = MIIP
                frmMain.Socket1.RemotePort = MIPORT
                frmMain.Socket1.Connect

        End If
       
 
End Select
Exit Sub
 
End Sub

Private Sub Image2_Click()
prgRun = False
End Sub

