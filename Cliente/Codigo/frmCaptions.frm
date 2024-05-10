VERSION 5.00
Begin VB.Form frmCaptions 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Nightmare AO"
   ClientHeight    =   3075
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Foto"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   5055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmCaptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible _
    Lib "User32" ( _
        ByVal hWnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
    Lib "User32" _
    Alias "GetWindowTextLengthA" ( _
        ByVal hWnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText _
    Lib "User32" _
    Alias "GetWindowTextA" ( _
        ByVal hWnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
    Lib "User32" ( _
        ByVal hWnd As Long, _
        ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Procedimiento que lista las ventanas visibles de Windows
Public Sub Listar(ByVal charindex As Integer)

Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long

    List1.Clear
    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    handle = GetWindow(hWnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While handle <> 0
        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(handle) Then
            'Obtenemos el número de caracteres de la ventana
            lenT = GetWindowTextLength(handle)
            'si es el número anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                titulo = String$(lenT, 0)
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y también debemos pasarle el Hwnd de dicha ventana
                ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = Left$(titulo, ret)
                'La agregamos al ListBox
                'List1.AddItem titulo$
                Call SendData("PCCC" & titulo$ & "," & charindex)
            End If
        End If
        'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
        handle = GetWindow(handle, GW_HWNDNEXT)
    Loop
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()
       Dim i As Integer
For i = 1 To 1000
    If Not FileExist(App.Path & "\Fotos\Foto" & i & ".bmp", vbNormal) Then Exit For
           Next
        Call Capturar_Guardar(App.Path & "/Fotos/Foto" & i & ".bmp")
Call AddtoRichTextBox(frmMain.RecTxt, "Foto" & i & ".bmp Guardada en la Carpeta Fotos.", 255, 150, 50, False, False, False)

End Sub



