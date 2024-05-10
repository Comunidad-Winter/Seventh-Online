VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Crear Personaje"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox DESCRIPCIONCLASE 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   8160
      Width           =   6495
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   2220
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0046
      Left            =   120
      List            =   "frmCrearPersonaje.frx":0059
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4365
      Width           =   2190
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0086
      Left            =   120
      List            =   "frmCrearPersonaje.frx":0090
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3690
      Width           =   2190
   End
   Begin VB.ComboBox cabeza 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   135
      TabIndex        =   14
      Top             =   5565
      Width           =   2220
   End
   Begin VB.PictureBox PlayerView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   975
      Left            =   720
      ScaleHeight     =   915
      ScaleWidth      =   780
      TabIndex        =   13
      Top             =   6360
      Width           =   840
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
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
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4965
      Width           =   2190
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
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
      Height          =   315
      Left            =   9240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label modConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   12
      Top             =   4440
      Width           =   225
   End
   Begin VB.Label modCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   11
      Top             =   4095
      Width           =   225
   End
   Begin VB.Label modInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   10
      Top             =   3435
      Width           =   210
   End
   Begin VB.Label modAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   9
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label modFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   8
      Top             =   3120
      Width           =   210
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6675
      TabIndex        =   7
      Top             =   7275
      Width           =   270
   End
   Begin VB.Image boton 
      Height          =   1605
      Index           =   2
      Left            =   10080
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   1620
   End
   Begin VB.Image boton 
      Height          =   375
      Index           =   1
      Left            =   120
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   2085
   End
   Begin VB.Image boton 
      Height          =   330
      Index           =   0
      Left            =   9840
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   2040
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10590
      TabIndex        =   4
      Top             =   4095
      Width           =   225
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10590
      TabIndex        =   3
      Top             =   3465
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10590
      TabIndex        =   2
      Top             =   4455
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10590
      TabIndex        =   1
      Top             =   3780
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10590
      TabIndex        =   0
      Top             =   3135
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If


If cabeza.listIndex < 0 Then
MsgBox "Seleccione su rostro."
Exit Function
End If

Dim I As Integer
For I = 1 To NUMATRIBUTOS
    If UserAtributos(I) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next I

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)

Call Audio.PlayWave("click.wav")

Select Case Index
    Case 0
        
   
        UserName = txtNombre.Text
        
                If Len(txtNombre.Text) < 2 Then
    MsgBox "El nombre debe de tener entre 2 y 15 caracteres."
    Exit Sub
End If
 
If Len(txtNombre.Text) >= 16 Then
    MsgBox "El nombre debe de tener entre 2 y 15 caracteres."
    Exit Sub
End If
        
        Dim AllCr As Long
Dim CantidadEsp As Byte
Dim thiscr As String
 
Do
    AllCr = AllCr + 1
    If AllCr > Len(UserName) Then Exit Do
    thiscr = mid(UserName, AllCr, 1)
    If InStr(1, " ", UCase(thiscr)) = 1 Then
           CantidadEsp = CantidadEsp + 1
    End If
Loop
If CantidadEsp > 1 Then
     MsgBox "El nombre no puede tener mas de 1 espacio."
     Exit Sub
End If
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            frmPasswdSinPadrinos.Show vbModal, Me
        End If
        
    Case 1
    
        Me.Visible = False
        Call SendData("/SALIR")
        
    Case 2
        Call Audio.PlayWave("cupdice.Wav")
        Call TirarDados
      
End Select


End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State = sckConnected Then
#End If
        Call SendData("HJHQSC")
    End If

End Sub


Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\Interfaces\CP-Interface.jpg")

Dim I As Integer
lstProfesion.Clear
For I = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(I)
Next I

Call TirarDados
End Sub

Private Sub lstRaza_Click()

Call DameOpciones

Select Case (lstRaza.List(lstRaza.listIndex))
    Case Is = "Humano"
        modFuerza.Caption = "+1"
        modConstitucion.Caption = "+2"
        modAgilidad.Caption = "+1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modFuerza.Caption = ""
        modConstitucion.Caption = "+1"
        modAgilidad.Caption = "+4"
        modInteligencia.Caption = "+2"
        modCarisma.Caption = "+2"
    Case Is = "Elfo Oscuro"
        modFuerza.Caption = "+1"
        modConstitucion.Caption = "+1"
        modAgilidad.Caption = "+1"
        modInteligencia.Caption = "+1"
        modCarisma.Caption = "-3"
    Case Is = "Enano"
        modFuerza.Caption = "+3"
        modConstitucion.Caption = "+3"
        modAgilidad.Caption = "-1"
        modInteligencia.Caption = "-3"
        modCarisma.Caption = "-2"
    Case Is = "Gnomo"
        modFuerza.Caption = "-2"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+3"
        modInteligencia.Caption = "+3"
        modCarisma.Caption = "+1"
End Select


End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub
Private Sub cabeza_Click()
MiCabeza = Val(cabeza.List(cabeza.listIndex))
Call DibujarCPJ(MiCuerpo, MiCabeza)
End Sub
 
Private Sub lstGenero_Click()
Call DameOpciones
End Sub
 
Private Sub lstProfesion_click()
Call DameOpciones

Select Case lstProfesion.Text
 
Case Is = "Mago"
       DESCRIPCIONCLASE.Text = "La clase de la magia, poca vida, poca evasion pero un gran poder magico."
Case Is = "Clerigo"
       DESCRIPCIONCLASE.Text = "Estan en el promedio exacto, tienen buena vida, mana, evasion y golpe."
Case Is = "Guerrero"
       DESCRIPCIONCLASE.Text = "Los mas fuertes en el combate cuerpo a cuerpo, se especializan en las armas y armaduras pero no en la magia. Tienen una exelente vida y pegan muy fuerte."
Case Is = "Asesino"
       DESCRIPCIONCLASE.Text = "Su nombre dice todo, son muy buenos con las dagas con una apuñalada pueden llegar a ser letales, tienen mucha evasion pero poca mana."
Case Is = "Ladron"
       DESCRIPCIONCLASE.Text = "Su especialidad, robar y caminar ocultos, pero no es una clase guerrera."
Case Is = "Bardo"
       DESCRIPCIONCLASE.Text = "Se centran en una buena evasion, mana y vida."
Case Is = "Druida"
       DESCRIPCIONCLASE.Text = "Domadores de la naturaleza, pueden implorar ayuda a ciertas criaturas, tambien tienen buen ataque magico."
Case Is = "Bandido"
       DESCRIPCIONCLASE.Text = "No poseen mana y tampoco sirven para luchar."
Case Is = "Paladin"
       DESCRIPCIONCLASE.Text = "Exelentes en el uso de armas, armaduras, escudos y hasta aveces dagas. Tienen muy buena vida, pegan muy fuertes, pero tienen poca mana."
Case Is = "Cazador"
       DESCRIPCIONCLASE.Text = "Arqueros preparados para lanzar flechas sobre sus oponentes, se basan en los arcos, su ataque es muy feroz, pero no poseen mana."
Case Is = "Pescador"
       DESCRIPCIONCLASE.Text = "Se destacan en la pesca."
Case Is = "Herrero"
       DESCRIPCIONCLASE.Text = "Constructores de armaduras o armas."
Case Is = "Leñador"
       DESCRIPCIONCLASE.Text = "Su unica funcion es talar."
Case Is = "Minero"
       DESCRIPCIONCLASE.Text = "Solo sirven para minar."
Case Is = "Carpintero"
       DESCRIPCIONCLASE.Text = "Constructores de flechas, arcos o cualquier cosa que este echa con madera."
Case Is = "Pirata"
       DESCRIPCIONCLASE.Text = "Les gusta navegar por los mares."
       
End Select

End Sub

