VERSION 5.00
Begin VB.Form frmCabezas 
   BorderStyle     =   0  'None
   Caption         =   "Eleccion de cabezas."
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   236
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Cerrar 
      Caption         =   "X"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.OptionButton Mujer 
      Caption         =   "Mujer"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.OptionButton Hombre 
      Caption         =   "Hombre"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   720
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton FechaDerecha 
      Caption         =   ">"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton FlechaIzquierda 
      Caption         =   "<"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Eleccion 
      Caption         =   "Enano"
      Height          =   375
      Index           =   6
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Eleccion 
      Caption         =   "Gnomo"
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Eleccion 
      Caption         =   "Elfo Oscuro"
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Eleccion 
      Caption         =   "Elfo"
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Eleccion 
      Caption         =   "Humano"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Cabeza 
      Caption         =   "Elejir cabeza."
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox Head 
      BackColor       =   &H00000000&
      Height          =   480
      Left            =   1560
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   2160
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Eleccion de Cabezas. By Midraks."
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmCabezas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeza_Click()
    If Actual <> 0 Then
        Call SendData("CAMBIOHEAD" & Actual)
    Else
        MsgBox "Debe elegir una cabeza.", , "Elegir cabezas. By Midraks"
    End If
End Sub

Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Eleccion_Click(Index As Integer)

        If Hombre.value = True Then
            Select Case Index
                Case 2
                    Actual = 1
                    MaxEleccion = 30
                    MinEleccion = 1
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
                Case 3
                    Actual = 101
                    MaxEleccion = 113
                    MinEleccion = 101
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
                Case 4
                    Actual = 202
                    MaxEleccion = 209
                    MinEleccion = 202
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
                Case 5
                    Actual = 401
                    MaxEleccion = 406
                    MinEleccion = 401
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
                Case 6
                    Actual = 301
                    MaxEleccion = 305
                    MinEleccion = 301
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
            End Select
        
        ElseIf Mujer.value = True Then
            Select Case Index
                Case 2
                    Actual = 70
                    MaxEleccion = 76
                    MinEleccion = 70
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
                Case 3
                    Actual = 170
                    MaxEleccion = 176
                    MinEleccion = 170
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
                Case 4
                    Actual = 270
                    MaxEleccion = 280
                    MinEleccion = 270
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
                Case 5
                    Actual = 470
                    MaxEleccion = 474
                    MinEleccion = 470
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
                Case 6
                    Actual = 370
                    MaxEleccion = 373
                    MinEleccion = 370
                    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
            End Select
        End If
        
End Sub

Private Sub FechaDerecha_Click()
    Actual = Actual + 1
    If Actual > MaxEleccion Then
        Actual = MaxEleccion
    ElseIf Actual < MinEleccion Then
        Actual = MinEleccion
    End If
    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
End Sub

Private Sub FlechaIzquierda_Click()
    Actual = Actual - 1
    If Actual > MaxEleccion Then
        Actual = MaxEleccion
    ElseIf Actual < MinEleccion Then
        Actual = MinEleccion
    End If
    Call engine.GrhRenderToHdc(HeadData(Actual).Head(3).grhindex, Head.hdc, 5, 5)
End Sub
