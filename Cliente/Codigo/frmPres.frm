VERSION 5.00
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Cargando"
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3400
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim puedo As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then If puedo Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = 800 * Screen.TwipsPerPixelX
    Me.Height = 600 * Screen.TwipsPerPixelY
    puedo = False
End Sub

Private Sub Timer1_Timer()

    Unload Me

End Sub
