VERSION 5.00
Begin VB.Form FrmPass 
   Caption         =   "               Cambiar Contraseña"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Text            =   "Repetir Contraseña Nueva:"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Text            =   "Contraseña Nueva:"
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Confirmar"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Height          =   3495
      Left            =   -240
      TabIndex        =   4
      Top             =   -360
      Width           =   4935
   End
End
Attribute VB_Name = "FrmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Text1.Text <> Text2.Text Then
    MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Sub
End If
If Len(Text1.Text) < 6 Or Len(Text1.Text) > 15 Then
MsgBox "Las contraseñas tienen que tenér un mínimo de 6 caracteres y un máximo de 15 cdaracteres"
Else
Call SendData("/passwd" & " " & Text1.Text)
MsgBox "||La password a sido cambiada exitosamente tu nueva clave es:" & Text1.Text
Unload Me
End If

End Sub
