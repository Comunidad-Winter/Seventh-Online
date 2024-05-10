VERSION 5.00
Begin VB.Form frmCuenta 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptionPJ 
      Caption         =   "Option1"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.OptionButton OptionPJ 
      Caption         =   "Option1"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.OptionButton OptionPJ 
      Caption         =   "Option1"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton OptionPJ 
      Caption         =   "Option1"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.OptionButton OptionPJ 
      Caption         =   "Option1"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.OptionButton OptionPJ 
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.OptionButton OptionPJ 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.OptionButton OptionPJ 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear Pejota"
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "frmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Boton_Click(Index As Integer)
Select Case Index
    Case 0
        'Boton Conectar
        If NamePJ(PJActive) = "No hay personaje." Then
            MsgBox "Debes seleccionar un personaje. Si no tienes personajes en tu cuenta, crea uno."
            Exit Sub
        End If
        NamePJLogued = NamePJ(PJActive)
        SendData ("OLOGIN" & NamePJLogued & "," & UserPassword)
        Exit Sub
       
    Case 1
        'Boton Crear PJ
        Me.Visible = False
        frmCrearPersonaje.Show
        Exit Sub
    End Select
End Sub
 
Private Sub OptionPJ_Click(Index As Integer)
PJActive = Index + 1
End Sub
