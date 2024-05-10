VERSION 5.00
Begin VB.Form Torneos 
   BackColor       =   &H00000000&
   Caption         =   "Torneo 1vs1 y 2vs2.       By Dylan.-"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   9600
      TabIndex        =   36
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   12240
      TabIndex        =   34
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   9840
      TabIndex        =   32
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Gana Torneo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   31
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   7800
      TabIndex        =   30
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Pierde"
      Height          =   375
      Left            =   4680
      TabIndex        =   29
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton Command24 
      Caption         =   "2 - 1 A Favor"
      Height          =   375
      Left            =   10320
      TabIndex        =   28
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "2 - 0 A Favor"
      Height          =   375
      Left            =   10320
      TabIndex        =   27
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "1 - 1"
      Height          =   375
      Left            =   10320
      TabIndex        =   26
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "1 - 0 A Favor"
      Height          =   375
      Left            =   10320
      TabIndex        =   25
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "1 - 1"
      Height          =   375
      Left            =   4680
      TabIndex        =   24
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Pierde"
      Height          =   375
      Left            =   10320
      TabIndex        =   23
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "2 - 1 A Favor"
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "2 - 0 A Favor"
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "1 - 0 A Favor"
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Duelean Ring 2 vs 2"
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   1080
      Width           =   9255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7200
      TabIndex        =   17
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   16
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command13 
      Caption         =   "2 - 1 A Favor"
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "2 - 1 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Pierde"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "2 - 0 A Favor"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "1 - 1"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1 - 0 A Favor"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Gana Torneo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pierde"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "2 - 0 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1 - 1"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1 - 0 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Duelean Ring 1 vs 1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Torneos 2 vs 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8040
      TabIndex        =   39
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Torneos 1vs1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1200
      TabIndex        =   38
      Top             =   120
      Width           =   2190
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9240
      TabIndex        =   37
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
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
      Height          =   255
      Left            =   11760
      TabIndex        =   35
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "VS"
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
      Height          =   255
      Left            =   9240
      TabIndex        =   33
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Y"
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
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "VS"
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
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Torneos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call SendData("/PELEA" & " " & Text1.Text & "@" & Text3.Text)

End Sub

Private Sub Command10_Click()
Call SendData("/RMSG Torneo> 2 A 0 A favor de" & " " & Text3.Text & "~")
Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

Private Sub Command11_Click()
Call SendData("/RMSG Torneo> Pierde:" & " " & Text3.Text & " Y " & Text1.Text & " Sigue de Ronda." & "~")
Call SendData("/TELEP" & " " & Text3.Text & " " & "1 50 50")
End Sub

Private Sub Command12_Click()
Call SendData("/RMSG Torneo> 2 A 1 A favor de" & " " & Text1.Text & "~")
Call SendData("/REVIVIR" & " " & Text3.Text)
End Sub

Private Sub Command13_Click()
Call SendData("/RMSG Torneo> 2 A 1 A favor de" & " " & Text3.Text & "~")
Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

Private Sub Command14_Click()
Call SendData("/VERSUS" & " " & Text2.Text & "@" & Text4.Text & "@" & Text7.Text & "@" & Text8.Text)
End Sub




Private Sub Command16_Click()
Call SendData("/RMSG Torneo 2 vs 2> 1 A 0 A favor de" & " " & Text2.Text & " Y " & Text4.Text & "~")
Call SendData("/REVIVIR" & " " & Text7.Text)
Call SendData("/REVIVIR" & " " & Text8.Text)
End Sub

Private Sub Command17_Click()
Call SendData("/RMSG Torneo 2 vs 2> 2 A 0 A favor de" & " " & Text2.Text & " Y " & Text4.Text & "~")
Call SendData("/REVIVIR" & " " & Text7.Text)
Call SendData("/REVIVIR" & " " & Text8.Text)
End Sub

Private Sub Command18_Click()
Call SendData("/RMSG Torneo 2 vs 2> 2 A 1 A favor de" & " " & Text2.Text & " Y " & Text4.Text & "~")
Call SendData("/REVIVIR" & " " & Text7.Text)
Call SendData("/REVIVIR" & " " & Text8.Text)
End Sub

Private Sub Command19_Click()
Call SendData("/RMSG Torneo 2 vs 2> Pierden:" & " " & Text7.Text & " Y " & Text8.Text & "~")
Call SendData("/TELEP" & " " & Text7.Text & " " & "1 50 50")
Call SendData("/TELEP" & " " & Text8.Text & " " & " 1 50 50")
End Sub



Private Sub Command20_Click()
Call SendData("/RMSG Torneo 2 vs 2> Lo empatan" & " " & Text2.Text & " Y " & Text4.Text & "~")
Call SendData("/REVIVIR" & " " & Text7.Text)
Call SendData("/REVIVIR" & " " & Text8.Text)
End Sub

Private Sub Command21_Click()
Call SendData("/RMSG Torneo 2 vs 2> 1 A 0 A favor de" & " " & Text7.Text & "Y " & Text8.Text & "~")
Call SendData("/REVIVIR" & " " & Text2.Text)
Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command22_Click()
Call SendData("/RMSG Torneo 2 vs 2> Lo empatan" & " " & Text7.Text & " Y " & Text8.Text & "~")
Call SendData("/REVIVIR" & " " & Text2.Text)
Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command23_Click()
Call SendData("/RMSG Torneo 2 vs 2> 2 A 0 A favor de" & " " & Text7.Text & " Y " & Text8.Text & "~")
Call SendData("/REVIVIR" & " " & Text2.Text)
Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command24_Click()
Call SendData("/RMSG Torneo 2 vs 2> 2 A 1 A favor de" & " " & Text7.Text & " Y " & Text8.Text & "~")
Call SendData("/REVIVIR" & " " & Text2.Text)
Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command25_Click()
Call SendData("/RMSG Torneo 2 vs 2> Pierden:" & " " & Text4.Text & " Y " & Text2.Text & "." & "~")
Call SendData("/TELEP" & " " & Text4.Text & " " & "1 50 50")
Call SendData("/TELEP" & " " & Text2.Text & " " & "1 50 50")
End Sub

Private Sub Command26_Click()
Call SendData("/RMSG Torneo 2 vs 2> Y el Ganador del Torneo es:" & " " & Text6.Text & " Y " & Text9.Text & "~")
Call SendData("/RMSG Gracias por Participar..! y Felicitaciones a" & " " & Text6.Text & " Y " & Text9.Text & "!!" & "~")
End Sub

Private Sub Command3_Click()
Call SendData("/RMSG Torneo> 1 A 0 A favor de" & " " & Text1.Text & "~")
Call SendData("/REVIVIR" & " " & Text3.Text)
End Sub
Private Sub Command4_Click()
Call SendData("/RMSG Torneo> Lo empata" & " " & Text1.Text & "~")
Call SendData("/REVIVIR" & " " & Text3.Text)
End Sub

Private Sub command5_Click()
Call SendData("/RMSG Torneo> 2 A 0 A favor de" & " " & Text1.Text & "~")
Call SendData("/REVIVIR" & " " & Text3.Text)
End Sub

Private Sub Command6_Click()
Call SendData("/RMSG Torneo> Pierde:" & " " & Text1.Text & " Y " & Text3.Text & " Sigue de Ronda." & "~")
Call SendData("/TELEP" & " " & Text1.Text & " " & "1 50 50")
End Sub

Private Sub Command7_Click()
Call SendData("/RMSG Torneo> Y el Ganador del Torneo es:" & " " & Text5.Text)
Call SendData("/RMSG Gracias por Participar..! y Felicitaciones a" & " " & Text5.Text & "!!" & "~")

End Sub

Private Sub Command8_Click()
Call SendData("/RMSG Torneo> 1 A 0 A favor de" & " " & Text3.Text & "~")
Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

Private Sub Command9_Click()
Call SendData("/RMSG Torneo> Lo empata" & " " & Text3.Text & "~")
Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

