VERSION 5.00
Begin VB.Form frmRetos 
   Caption         =   "Retos."
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butRetar 
      Caption         =   "Retar"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   2880
      Width           =   3255
   End
   Begin VB.CommandButton butVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtPareja 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox txtCont 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtOro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.OptionButton opc2v2 
      Caption         =   "Jugar 2vs2"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.OptionButton opc1V1 
      Caption         =   "Jugar 1vs1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Name Pareja"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Name Contrincante"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad de oro a apostar:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butRetar_Click()
If Len(Trim(txtOro.Text)) = "" Then
MsgBox "Pon la cantidad de oro"
Exit Sub
End If
If Val(txtOro.Text) < 5000 Then
MsgBox "Tu apuesta debe ser mayor a 5000."
Exit Sub
End If
If opc1V1.value = True Then
    If Len(Trim(txtCont.Text)) = "" Then
    MsgBox "Pon el name del usuario"
    Exit Sub
    End If
End If
SendData "/RETAR " & txtCont.Text & "," & txtOro.Text
If opc2v2.value = True Then
    If Len(Trim(txtPareja.Text)) = "" Then
    MsgBox "Pon el nombre de tu pareja"
    Exit Sub
    End If
    SendData "/PAREJA " & txtCont.Text & "," & txtOro.Text
End If
End Sub

Private Sub butVolver_Click()
Unload Me
End Sub

Private Sub opc1V1_Click()
txtCont.Locked = False
txtContpareja.Locked = True
txtPareja.Locked = True
End Sub

Private Sub opc2v2_Click()
txtCont.Locked = True
txtContpareja.Locked = False
txtPareja.Locked = False
End Sub
