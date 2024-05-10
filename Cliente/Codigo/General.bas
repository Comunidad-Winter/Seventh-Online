Attribute VB_Name = "Mod_General"

Option Explicit

Public LastTick As Long
Public ThisTick As Long
Public OldTick As Long
Public OldQuery As Long
Public FirstTimeChit As Long
Public Trys As Integer


Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public mutexHID As Long

Dim CQC As New cTickCount

Public Declare Sub keybd_event Lib "user32" ( _
ByVal bVk As Byte, _
ByVal bScan As Byte, _
ByVal dwFlags As Long, _
ByVal dwExtraInfo As Long)
 
Public Const VK_SNAPSHOT = &H2C

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
                 Private Const LWA_ALPHA = &H2
Public Declare Function ReleaseCapture Lib "user32" () As Long

Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
Const RGN_OR = 2


Public Declare Function SendMessages Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    
    Public Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
    End Type
    Public Const WM_COPYDATA = &H4A
    
    
Private Const ERROR_ALREADY_EXISTS = 183&

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


Public bK As Long
Public bRK As Long

Public bFogata As Boolean

Public lFrameTimer As Long

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal r As String)
 
 
Public Function MD5String(p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, T As Long
    r = Space(32)
    T = Len(p)
    MDStringFix p, T, r
    MD5String = r
End Function
Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function


Public Function DirGraficos() As String
    DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function

Public Function SumaDigitos(ByVal Numero As Integer) As Integer
    'Suma digitos
    Do
        SumaDigitos = SumaDigitos + (Numero Mod 10)
        Numero = Numero \ 10
    Loop While (Numero > 0)
End Function

Public Function SumaDigitosMenos(ByVal Numero As Integer) As Integer
    'Suma digitos, y resta el total de dígitos
    Do
        SumaDigitosMenos = SumaDigitosMenos + (Numero Mod 10) - 1
        Numero = Numero \ 10
    Loop While (Numero > 0)
End Function

Public Function Complex(ByVal Numero As Integer) As Integer
    If Numero Mod 2 <> 0 Then
        Complex = Numero * SumaDigitos(Numero)
    Else
        Complex = Numero * SumaDigitosMenos(Numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal Numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(Numero)
    AuxInteger2 = SumaDigitosMenos(Numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.Path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorH:

    Versiones(1) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.Path & "\init\" & "versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
    Dim archivoC As String
    
    archivoC = App.Path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim I As Long
    
    For I = 0 To 47 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(I).r = CByte(GetVar(archivoC, CStr(I), "R"))
        ColoresPJ(I).G = CByte(GetVar(archivoC, CStr(I), "G"))
        ColoresPJ(I).b = CByte(GetVar(archivoC, CStr(I), "B"))
    Next I
    
ColoresPJ(50).r = 255
ColoresPJ(50).G = 0
ColoresPJ(50).b = 0
ColoresPJ(49).r = 60
ColoresPJ(49).G = 94
ColoresPJ(49).b = 255
ColoresPJ(48).r = 120
ColoresPJ(48).G = 120
ColoresPJ(48).b = 120
End Sub
Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.Path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    If ConsolaActivada = True Then
        With RichTextBox
            If (Len(.Text)) > 10000 Then .Text = ""
            
            .SelStart = Len(RichTextBox.Text)
            .SelLength = 0
            
            .SelBold = Bold
            .SelItalic = Italic
            
            If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
            
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
        End With
    Else:
        Exit Sub
    End If
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).Active = 1 Then
            MapData(charlist(loopc).Pos.x, charlist(loopc).Pos.Y).charindex = loopc
        End If
    Next loopc
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim I As Long
    
    cad = LCase$(cad)
    
    For I = 1 To Len(cad)
        car = Asc(mid$(cad, I, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next I
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function


Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()

    Connected = True
    
    Unload frmConnect
    
    StartAntiSH
    
    frmMain.Label8.Caption = UserName
    
    Call SetMusicInfo("Jugando SeventhAO [" & UserName & "] - [www.seventh-ao.com.ar]", "Games", "{1}{0}")

    frmMain.Visible = True
    AddtoRichTextBox frmMain.RecTxt, "Bienvenido " & UserName & " a Seventh Online, suerte y disfruta.", 128, 128, 0, True, False, False
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.x, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.x + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.x, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.x - 1, UserPos.Y)
    End Select
    
    If UserMeditar Then UserMeditar = False

If UserParalizado Then
     If charlist(UserCharIndex).Heading <> Direccion Then
        Call SendData("CHEA" & Direccion)
     End If
     Exit Sub
    End If
    
    If LegalOk Then
        Call SendData("M" & Direccion)
        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call SendData("CHEA" & Direccion)
        End If
    End If
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub

Sub CheckKeys() 'Stand
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
                Call MoveTo(NORTH)
                Call MovemosUserMap
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.Y & ")"
                Exit Sub
            End If
       
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
                Call MoveTo(EAST)
                Call MovemosUserMap
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.Y & ")"
                Exit Sub
            End If
       
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
                Call MoveTo(SOUTH)
                Call MovemosUserMap
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.Y & ")"
                Exit Sub
            End If
       
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
                Call MoveTo(WEST)
                Call MovemosUserMap
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.Y & ")"
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            If kp Then Call RandomMove
            Call MovemosUserMap
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
            If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                End If
            frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.Y & ")"
        End If
    End If
End Sub
'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim loopc As Long
    Dim Y As Long
    Dim x As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    
    Open DirMapas & "Mapa" & Map & ".map" For Binary As #1
    Seek #1, 1
            
    'map Header
    Get #1, , MapInfo.MapVersion
    Get #1, , MiCabecera
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            Get #1, , ByFlags
            
            MapData(x, Y).Blocked = (ByFlags And 1)
            
            Get #1, , MapData(x, Y).Graphic(1).GrhIndex
            InitGrh MapData(x, Y).Graphic(1), MapData(x, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get #1, , MapData(x, Y).Graphic(2).GrhIndex
                InitGrh MapData(x, Y).Graphic(2), MapData(x, Y).Graphic(2).GrhIndex
            Else
                MapData(x, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get #1, , MapData(x, Y).Graphic(3).GrhIndex
                InitGrh MapData(x, Y).Graphic(3), MapData(x, Y).Graphic(3).GrhIndex
            Else
                MapData(x, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get #1, , MapData(x, Y).Graphic(4).GrhIndex
                InitGrh MapData(x, Y).Graphic(4), MapData(x, Y).Graphic(4).GrhIndex
            Else
                MapData(x, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get #1, , MapData(x, Y).Trigger
            Else
                MapData(x, Y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(x, Y).charindex > 0 Then
                Call EraseChar(MapData(x, Y).charindex)
            End If
            
            'Erase OBJs
            MapData(x, Y).ObjGrh.GrhIndex = 0
            MapData(x, Y).ObjName = ""
        Next x
    Next Y
    
    Close #1
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
    Call DibujarMiniMapa
    Call MovemosUserMapa
End Sub

'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim I As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For I = 1 To Len(Text)
        CurChar = mid$(Text, I, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = I
        End If
    Next I
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid$(Text, LastPos + 1)
    End If
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.Path & "\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
    Dim sa As SECURITY_ATTRIBUTES
   
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)
    End With
   
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function
 
Public Function FindPreviousInstance() As Boolean
    If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
        FindPreviousInstance = False
    Else
        FindPreviousInstance = True
    End If
End Function
 
Public Sub ReleaseInstance()
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub


Sub Main()

HDSerial = GetDriveSerialNumber

SetResolution

On Error Resume Next


    Call WriteClientVer
    Call LeerLineaComandos
    
    If App.PrevInstance Then
        Call MsgBox("SeventhAO ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
       
Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 2) As Integer

    ChDrive App.Path
    ChDir App.Path

    'Cargamos el archivo de configuracion inicial
    If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
    
    
    If FileExist(App.Path & "\init\ao.dat", vbArchive) Then
        Call LoadClientSetup
        
        If ClientSetup.bDinamic Then
            Set SurfaceDB = New clsSurfaceManDyn
        Else
            Set SurfaceDB = New clsSurfaceManStatic
        End If
    Else
        'Por default usamos el dinámico
        Set SurfaceDB = New clsSurfaceManDyn
    End If
    
    
    tipf = Config_Inicio.tip
    
    frmCargando.Show
    frmCargando.Refresh
    
    frmCargando.Barra.Width = frmCargando.Barra.Width + 100
    
    AddtoRichTextBox frmCargando.status, "Buscando servidores....", 0, 0, 0, 0, 0, 1

#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

    AddtoRichTextBox frmCargando.status, "Iniciando constantes...", 0, 0, 0, 0, 0, 1
    
    Call InicializarNombres
    
    
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1
    
    frmCargando.Barra.Width = frmCargando.Barra.Width + 100
    
    AddtoRichTextBox frmCargando.status, "Cargando Sonidos....", 0, 0, 0, 0, 0, 1
    AddtoRichTextBox frmCargando.status, "Hecho", , , , 1

Dim loopc As Integer

lastTime = GetTickCount

    Call InitTileEngine(frmMain.hwnd, 152, 7, 32, 32, 13, 17, 9)
    
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra....")
    
    frmCargando.Barra.Width = 548
    Call CargarAnimsExtra

UserMap = 1

    Call CargarAnimArmas
    Call CargarAnimEscudos
        Call CargarVersiones
    Call CargarColores

    AddtoRichTextBox frmCargando.status, "                    ¡Bienvenido a Argentum Online!", , , , 1
    'Activar/Desactivar Consola
    ConsolaActivada = True
    
    Unload frmCargando
    
    'Inicializamos el sonido
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound....", 0, 0, 0, 0, 0, True)
    Call Audio.Initialize(DirectX, frmMain.hwnd, App.Path & "\" & Config_Inicio.DirSonidos & "\", App.Path & "\" & Config_Inicio.DirMusica & "\")
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1, , False)
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectDraw, frmMain.picInv)
    
    If Musica Then
        Call Audio.PlayMIDI("174.mid")
    End If

    'frmPres.Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Cargando.jpg")
    'frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    
    frmConnect.Visible = True

'TODO : Esto va en Engine Initialization
    MainViewRect.Left = MainViewLeft
    MainViewRect.Top = MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
'TODO : Esto va en Engine Initialization
    MainDestRect.Left = TilePixelWidth * TileBufferSize - TilePixelWidth
    MainDestRect.Top = TilePixelHeight * TileBufferSize - TilePixelHeight
    MainDestRect.Right = MainDestRect.Left + MainViewWidth
    MainDestRect.Bottom = MainDestRect.Top + MainViewHeight
    
    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame
             Call speedHackCheck
            'Play ambient sounds
            Call RenderSounds
        End If
        
'TODO : Porque el pausado de 20 ms???
        If GetTickCount - lastTime > 20 Then
            If Not pausa And frmMain.Visible And Not frmForo.Visible And Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmBancoObj.Visible Then
                CheckKeys
                lastTime = GetTickCount
            End If
        End If
        
        While (GetTickCount - lFrameTimer) \ 56 < FramesPerSecCounter
            Sleep 5
        Wend
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            FramesPerSec = FramesPerSecCounter
            
            If FPSFLAG Then frmMain.Caption = FramesPerSec
            
            FramesPerSecCounter = 0
            lFrameTimer = GetTickCount
        End If
        
'TODO : Sería mejor comparar el tiempo desde la última vez que se hizo hasta el actual SOLO cuando se precisa. Además evitás el corte de intervalos con 2 golpes seguidos.
        'Sistema de timers renovado:
        esttick = GetTickCount
        For loopc = 1 To UBound(timers)
            'Timer de trabajo
            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False
            End If
            'timer de attaque (77)
            If timers(2) >= tAt Then
                timers(2) = 0
                UserCanAttack = 1
                UserPuedeRefrescar = True
            End If
        Next loopc
        ulttick = GetTickCount
        
#If SeguridadAlkon Then
        Call CheckSecurity
#End If
        
        DoEvents
    Loop

    EngineRun = False
    frmCargando.Show
    AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
    DeinitTileEngine

'TODO : Esto debería ir en otro lado como al cambair a esta res


    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing

    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    
#If SeguridadAlkon Then
    DeinitSecurity
#End If
End

ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal x As Integer, ByVal Y As Integer) As Boolean

    HayAgua = MapData(x, Y).Graphic(1).GrhIndex >= 1505 And _
                MapData(x, Y).Graphic(1).GrhIndex <= 1520 And _
                MapData(x, Y).Graphic(2).GrhIndex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim I As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For I = LBound(T) To UBound(T)
        Select Case UCase$(T(I))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next I
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open App.Path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle
    
    Musica = Not ClientSetup.bNoMusic
    Sound = Not ClientSetup.bNoSound
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(1) = "Ullathorpe"
    Ciudades(2) = "Nix"
    Ciudades(3) = "Banderbill"

    CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
    CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
    CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Ladron"
    ListaClases(6) = "Bardo"
    ListaClases(7) = "Druida"
    ListaClases(8) = "Bandido"
    ListaClases(9) = "Paladin"
    ListaClases(10) = "Cazador"
    ListaClases(11) = "Pescador"
    ListaClases(12) = "Herrero"
    ListaClases(13) = "Leñador"
    ListaClases(14) = "Minero"
    ListaClases(15) = "Carpintero"
    ListaClases(16) = "Pirata"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.Meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"
    SkillsNames(Skills.Equitacion) = "Equitacion"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub
Public Sub SetMusicInfo(ByRef r_sArtist As String, ByRef r_sAlbum As String, ByRef r_sTitle As String, Optional ByRef r_sWMContentID As String = vbNullString, Optional ByRef r_sFormat As String = "{0} - {1}", Optional ByRef r_bShow As Boolean = True)
 
Dim udtData As COPYDATASTRUCT
Dim sBuffer As String
Dim hMSGRUI As Long
 
'Total length can Not be longer Then 256 characters!
'Any longer will simply be ignored by Messenger.
sBuffer = "\0Games\0" & Abs(r_bShow) & "\0" & r_sFormat & "\0" & r_sArtist & "\0" & r_sTitle & "\0" & r_sAlbum & "\0" & r_sWMContentID & "\0" & vbNullChar
 
udtData.dwData = &H547
udtData.lpData = StrPtr(sBuffer)
udtData.cbData = LenB(sBuffer)
 
Do
hMSGRUI = FindWindowEx(0&, hMSGRUI, "MsnMsgrUIManager", vbNullString)
 
If (hMSGRUI > 0) Then
Call SendMessages(hMSGRUI, WM_COPYDATA, 0, VarPtr(udtData))
End If
 
Loop Until (hMSGRUI = 0)
 
End Sub

Public Sub HookSurfaceHwnd(pic As PictureBox)
    Call ReleaseCapture
    Call SendMessage(pic.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
 

Public Function PonerPuntos(Numero As Long) As String
Dim I As Integer
Dim Cifra As String
 
Cifra = Str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For I = 0 To 4
    If Len(Cifra) - 3 * I >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * I), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * I), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * I > 0 Then
            PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * I) & "." & PonerPuntos
        End If
        Exit For
    End If
Next
 
PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
End Function

Public Function Transparencia(ByVal hwnd As Long, _
                                      Valor As Integer) As Long
  
Dim Transparencias As Long
   Transparencias = GetWindowLong(hwnd, GWL_EXSTYLE)
   Transparencias = Transparencias Or WS_EX_LAYERED
     
   SetWindowLong hwnd, GWL_EXSTYLE, Transparencias
   SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
  
End Function

Public Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As Long
'***************************************************
'Author: Nahuel Casas (Zagen)
'Last Modify Date: 07/12/2009
' 07/12/2009: Zagen - Convertì las funciones, en formulas mas fàciles de modificar.
'***************************************************
    On Error Resume Next
          Dim fso As Object, Drv As Object, DriveSerial As Long
         
          'Creamos el objeto FileSystemObject.
          Set fso = CreateObject("Scripting.FileSystemObject")
         
          'Asignamos el driver principal.
          If DriveLetter <> "" Then
              Set Drv = fso.GetDrive(DriveLetter)
          Else
              Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))
          End If
     
          With Drv
              If .IsReady Then
                  DriveSerial = Abs(.SerialNumber)
              Else    '"Si el driver no està como para empezar ..."
                  DriveSerial = -1
              End If
          End With
         
          'Borramos y limpiamos.
          Set Drv = Nothing
          Set fso = Nothing
    'Seteamos :)
    GetDriveSerialNumber = DriveSerial
         
End Function

Public Sub Capturar_Guardar(Path As String)
Clipboard.Clear
keybd_event VK_SNAPSHOT, 1, 0, 0
DoEvents
    If Clipboard.GetFormat(vbCFBitmap) Then
            SavePicture Clipboard.GetData(vbCFBitmap), Path
            'MsgBox " Captura generada en: " & Path, vbInformation
    'Picture1.Picture = Clipboard.GetData(vbCFBitmap)
    Else
    MsgBox " Error ", vbCritical
    End If
End Sub
Public Sub LogError(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\errores.log" For Append As #nfile
Print #nfile, desc
Close #nfile
End Sub

Public Sub LogCustom(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\custom.log" For Append As #nfile
Print #nfile, Now & " " & desc
Close #nfile
End Sub

  Public Function asdasjfjlkawqwr(strText As String, ByVal strPwd As String)
        Dim I As Integer, C As Integer
        Dim strBuff As String
    #If Not CASE_SENSITIVE_PASSWORD Then
  
        strPwd = UCase$(strPwd)
    #End If
    
        If Len(strPwd) Then
            For I = 1 To Len(strText)
                C = Asc(mid$(strText, I, 1))
                C = C + Asc(mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
                strBuff = strBuff & Chr$(C And &HFF)
            Next I
        Else
            strBuff = strText
        End If
        asdasjfjlkawqwr = strBuff
    End Function
     
     
    Private Function Seventhasdlwqe(x As Integer) As String
        If x > 9 Then
            Seventhasdlwqe = Chr(x + 30)
        Else
            Seventhasdlwqe = CStr(x)
        End If
    End Function
     

    Public Function Encriptar(DataValue As Variant) As Variant
           
        Dim x As Long
        Dim temp As String
        Dim TempNum As Integer
        Dim TempChar As String
        Dim TempChar2 As String
           
        For x = 1 To Len(DataValue)
            TempChar2 = mid(DataValue, x, 1)
            TempNum = Int(Asc(TempChar2) / 16)
               
            If ((TempNum * 16) < Asc(TempChar2)) Then
                     
                TempChar = Seventhasdlwqe(Asc(TempChar2) - (TempNum * 16))
                temp = temp & Seventhasdlwqe(TempNum) & TempChar
            Else
                temp = temp & Seventhasdlwqe(TempNum) & "0"
               
            End If
        Next x
           
           
        Encriptar = temp
    End Function
     
    
    Public Function Seventhqwedvggjfgnb(strText As String, ByVal strPwd As String)
        Dim I As Integer, C As Integer
        Dim strBuff As String
     
    #If Not CASE_SENSITIVE_PASSWORD Then
     
        strPwd = UCase$(strPwd)
     
    #End If
     
    
        If Len(strPwd) Then
            For I = 1 To Len(strText)
                C = Asc(mid$(strText, I, 1))
                C = C - Asc(mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
                strBuff = strBuff & Chr$(C And &HFF)
            Next I
        Else
            strBuff = strText
        End If
        Seventhqwedvggjfgnb = strBuff
    End Function
     
     
    Private Function Seventhwqedfdhbvczsdf(x As String) As Integer
           
        Dim x1 As String
        Dim x2 As String
        Dim temp As Integer
           
        x1 = mid(x, 1, 1)
        x2 = mid(x, 2, 1)
           
        If IsNumeric(x1) Then
            temp = 16 * Int(x1)
        Else
            temp = (Asc(x1) - 30) * 16
        End If
           
        If IsNumeric(x2) Then
            temp = temp + Int(x2)
        Else
            temp = temp + (Asc(x2) - 30)
        End If
           
        ' retorno
        Seventhwqedfdhbvczsdf = temp
           
    End Function
     
   
    Function DameDameX(DataValue As Variant) As Variant
           
        Dim x As Long
        Dim temp As String
        Dim HexByte As String
           
        For x = 1 To Len(DataValue) Step 2
               
            HexByte = mid(DataValue, x, 2)
            temp = temp & Chr(Seventhwqedfdhbvczsdf(HexByte))
               
        Next x
        ' retorno
        DameDameX = temp
           
    End Function

Public Function GetTick() As Long
GetTick = CLng(CQC.GetElapsedTime(False) * 1000)
End Function

