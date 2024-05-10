Attribute VB_Name = "Mod_TileEngine"


Option Explicit

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    C       O       N       S      T
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    T       I      P      O      S
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Encabezado bmp
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    x As Integer
    Y As Integer
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
    Active As Boolean
    MiniMap_color As Long
End Type


'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
    'ANIM ATAK ESCUDO
     ShieldAttack As Byte
End Type


'Lista de cuerpos
Public Type FxData
    Fx As Grh
    OffsetX As Long
    OffsetY As Long
End Type

'Apariencia del personaje
Public Type Char
    Aura_Index As Integer
    Aura As Grh
    Active As Byte
    Heading As Byte ' As E_Heading ?
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    Fx As Integer
    FxLoopTimes As Integer
    EsStatus As Byte
    Criminal As Byte
    
    Nombre As String
    
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh
    ObjName As String
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type

 
Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7
 
Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public SecundaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7


Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public lastTime As Long 'Para controlar la velocidad


'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh 'Animaciones publicas
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Usuarios?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'
'epa ;)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿API?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

#If ConAlfaB Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

#End If

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Sub CargarCabezas()
On Error Resume Next
Dim n As Integer, I As Integer, Numheads As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , Numheads

'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For I = 1 To Numheads
    Get #n, , Miscabezas(I)
    InitGrh HeadData(I).Head(1), Miscabezas(I).Head(1), 0
    InitGrh HeadData(I).Head(2), Miscabezas(I).Head(2), 0
    InitGrh HeadData(I).Head(3), Miscabezas(I).Head(3), 0
    InitGrh HeadData(I).Head(4), Miscabezas(I).Head(4), 0
Next I

Close #n

End Sub

Sub CargarCascos()
On Error Resume Next
Dim n As Integer, I As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For I = 1 To NumCascos
    Get #n, , Miscabezas(I)
    InitGrh CascoAnimData(I).Head(1), Miscabezas(I).Head(1), 0
    InitGrh CascoAnimData(I).Head(2), Miscabezas(I).Head(2), 0
    InitGrh CascoAnimData(I).Head(3), Miscabezas(I).Head(3), 0
    InitGrh CascoAnimData(I).Head(4), Miscabezas(I).Head(4), 0
Next I

Close #n

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim n As Integer, I As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

n = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For I = 1 To NumCuerpos
    Get #n, , MisCuerpos(I)
    InitGrh BodyData(I).Walk(1), MisCuerpos(I).Body(1), 0
    InitGrh BodyData(I).Walk(2), MisCuerpos(I).Body(2), 0
    InitGrh BodyData(I).Walk(3), MisCuerpos(I).Body(3), 0
    InitGrh BodyData(I).Walk(4), MisCuerpos(I).Body(4), 0
    BodyData(I).HeadOffset.x = MisCuerpos(I).HeadOffsetX
    BodyData(I).HeadOffset.Y = MisCuerpos(I).HeadOffsetY
Next I

Close #n

End Sub
Sub CargarFxs()
On Error Resume Next
Dim n As Integer, I As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

n = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumFxs

'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For I = 1 To NumFxs
    Get #n, , MisFxs(I)
    Call InitGrh(FxData(I).Fx, MisFxs(I).Animacion, 1)
    FxData(I).OffsetX = MisFxs(I).OffsetX
    FxData(I).OffsetY = MisFxs(I).OffsetY
Next I

Close #n

End Sub

Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.x + CX
tY = UserPos.Y + CY

End Sub






Sub MakeChar(ByVal charindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

'Apuntamos al ultimo Char
If charindex > LastChar Then LastChar = charindex

NumChars = NumChars + 1

If Arma = 0 Then Arma = 2
If Escudo = 0 Then Escudo = 2
If Casco = 0 Then Casco = 2

charlist(charindex).iHead = Head
charlist(charindex).iBody = Body
charlist(charindex).Head = HeadData(Head)
charlist(charindex).Body = BodyData(Body)
charlist(charindex).Arma = WeaponAnimData(Arma)
'[ANIM ATAK]
charlist(charindex).Arma.WeaponAttack = 0

charlist(charindex).Escudo = ShieldAnimData(Escudo)
charlist(charindex).Casco = CascoAnimData(Casco)

charlist(charindex).Heading = Heading

'Reset moving stats
charlist(charindex).Moving = 0
charlist(charindex).MoveOffset.x = 0
charlist(charindex).MoveOffset.Y = 0

'Update position
charlist(charindex).Pos.x = x
charlist(charindex).Pos.Y = Y

'Make active
charlist(charindex).Active = 1

'Plot on map
MapData(x, Y).charindex = charindex

End Sub

Sub ResetCharInfo(ByVal charindex As Integer)

    charlist(charindex).EsStatus = 0
 
    charlist(charindex).Active = 0
    charlist(charindex).Criminal = 0
    charlist(charindex).Fx = 0
    charlist(charindex).FxLoopTimes = 0
    charlist(charindex).invisible = False

#If SeguridadAlkon Then
    Call MI(CualMI).ResetInvisible(charindex)
#End If

    charlist(charindex).Moving = 0
    charlist(charindex).muerto = False
    charlist(charindex).Nombre = ""
    charlist(charindex).pie = False
    charlist(charindex).Pos.x = 0
    charlist(charindex).Pos.Y = 0
    charlist(charindex).UsandoArma = False

End Sub


Sub EraseChar(ByVal charindex As Integer)
On Error Resume Next

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

charlist(charindex).Active = 0

'Update lastchar
If charindex = LastChar Then
    Do Until charlist(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.Y).charindex = 0

Call ResetCharInfo(charindex)

'Mithrandir reparando
Dialogos.QuitarDialogo (charindex)
'Mithrandir reparando

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1
'[CODE 000]:MatuX
'
'  La linea generaba un error en la IDE, (no ocurría debido al
' on error)
'
'   Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
'[END]'

End Sub

Sub MoveCharbyHead(ByVal charindex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim x As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

x = charlist(charindex).Pos.x
Y = charlist(charindex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        addY = -1

    Case E_Heading.EAST
        addX = 1

    Case E_Heading.SOUTH
        addY = 1
    
    Case E_Heading.WEST
        addX = -1
        
End Select

nX = x + addX
nY = Y + addY

MapData(nX, nY).charindex = charindex
charlist(charindex).Pos.x = nX
charlist(charindex).Pos.Y = nY
MapData(x, Y).charindex = 0

charlist(charindex).MoveOffset.x = -1 * (TilePixelWidth * addX)
charlist(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

charlist(charindex).Moving = 1
charlist(charindex).Heading = nHeading

If UserEstado <> 1 Then Call DoPasosFx(charindex)

'areas viejos
If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Debug.Print UserCharIndex
    Call EraseChar(charindex)
End If

End Sub

Public Sub DoFogataFx()
If Sound Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata()
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim x As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For x = UserPos.x - MinXBorder + 1 To UserPos.x + MinXBorder - 1
            
            If MapData(x, Y).charindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next x
Next Y

EstaPCarea = False

End Function


Sub DoPasosFx(ByVal charindex As Integer)
Static pie As Boolean

If Not Sound Then Exit Sub

If Not UserNavegando Then
    If Not charlist(charindex).muerto And EstaPCarea(charindex) Then
        charlist(charindex).pie = Not charlist(charindex).pie
        If charlist(charindex).pie Then
            Call Audio.PlayWave("23.wav")
        Else
            Call Audio.PlayWave("24.wav")
        End If
    End If
Else
    Call Audio.PlayWave("50.wav")
End If

End Sub


Sub MoveCharbyPos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)
If charindex = UserCharIndex And UserParalizado Then Exit Sub


On Error Resume Next

Dim x As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As E_Heading



x = charlist(charindex).Pos.x
Y = charlist(charindex).Pos.Y

MapData(x, Y).charindex = 0

addX = nX - x
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = E_Heading.EAST
End If

If Sgn(addX) = -1 Then
    nHeading = E_Heading.WEST
End If

If Sgn(addY) = -1 Then
    nHeading = E_Heading.NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = E_Heading.SOUTH
End If

MapData(nX, nY).charindex = charindex


charlist(charindex).Pos.x = nX
charlist(charindex).Pos.Y = nY

charlist(charindex).MoveOffset.x = -1 * (TilePixelWidth * addX)
charlist(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

charlist(charindex).Moving = 1
charlist(charindex).Heading = nHeading

'parche para que no medite cuando camina
Dim fxCh As Integer
fxCh = charlist(charindex).Fx
If fxCh = FxMeditar.CHICO Or fxCh = FxMeditar.GRANDE Or fxCh = FxMeditar.MEDIANO Or fxCh = FxMeditar.XGRANDE Or fxCh = FxMeditar.CIUDA Or fxCh = FxMeditar.CRIMI Or fxCh = FxMeditar.TRANSFO Then
    charlist(charindex).Fx = 0
    charlist(charindex).FxLoopTimes = 0
End If

If Not EstaPCarea(charindex) Then Dialogos.QuitarDialogo (charindex)

If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(charindex)
End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
   
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
   
        Case E_Heading.EAST
            x = 1
   
        Case E_Heading.SOUTH
            Y = 1
       
        Case E_Heading.WEST
            x = -1
           
    End Select
   
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.Y + Y
 
    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
       
        bTecho = IIf(MapData(UserPos.x, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.x - 8 To UserPos.x + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopc As Integer
Dim Dale As Boolean

loopc = 1
Do While charlist(loopc).Active And Dale
    loopc = loopc + 1
    Dale = (loopc <= UBound(charlist))
Loop

NextOpenChar = loopc

End Function


Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer




'Resize arrays
ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

'Open files
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
GrhData(Grh).Active = True

        
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1
Dim Count As Long
 
Open IniPath & "minimap.dat" For Binary As #1
    Seek #1, 1
    For Count = 1 To 15000
        If GrhData(Count).Active Then
            Get #1, , GrhData(Count).MiniMap_color
        End If
    Next Count
Close #1


Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Limites del mapa
If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    'Tile Bloqueado?
    If MapData(x, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(x, Y).charindex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(x, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(x, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function




Function InMapLegalBounds(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If x < XMinMapSize Or x > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If
'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
'Center Grh over X,Y pos
If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
Surface.BltFast x, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, Grh As Integer, ByVal x As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = x
    .Top = Y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.Surface(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

'Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[CODE 000]:MatuX
    Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With


Surface.BltFast x, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub

#If ConAlfaB = 1 Then
    Sub DDrawTransGrhtoSurfaceAlpha(Surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then

                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If

                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(x < 0, Abs(x), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set Src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)

Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = x
    .Top = Y
    .Right = x + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
    
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast x, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(x + x, Y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)

Surface.Unlock rDest
DstLock = False
Src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub
#End If 'ConAlfaB = 1

Sub DrawBackBufferSurface()
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hwnd As Long, hDC As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
    If Grh <= 0 Then Exit Sub
    
    SecundaryClipper.SetHWnd hwnd
    SurfaceDB.Surface(GrhData(Grh).FileNum).BltToDC hDC, SourceRect, destRect
End Sub

Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
On Error Resume Next


If UserCiego Then Exit Sub

Dim Y        As Integer 'Keeps track of where on map we are
Dim x        As Integer 'Keeps track of where on map we are
Dim minY     As Integer 'Start Y pos on current map
Dim maxY     As Integer 'End Y pos on current map
Dim minX     As Integer 'Start X pos on current map
Dim maxX     As Integer 'End X pos on current map
Dim ScreenX  As Integer 'Keeps track of where to place tile on screen
Dim ScreenY  As Integer 'Keeps track of where to place tile on screen
Dim Moved    As Byte
Dim Grh      As Grh     'Temp Grh for show tile and blocked
Dim TempChar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer 'Usado en el Layer de Chars
Dim iPPy     As Integer 'Usado en el Layer de Chars
Dim rSourceRect      As RECT    'Usado en el Layer 1
Dim iGrhIndex        As Integer 'Usado en el Layer 1
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim nX As Integer
Dim nY As Integer

'Figure out Ends and Starts of screen
' Hardcodeado para speed!
minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)


'Draw floor layer
ScreenY = 8
For Y = (minY + 8) To maxY - 8
    ScreenX = 8
    For x = minX + 8 To maxX - 8
        If x > 100 Or Y < 1 Then Exit For
        'Layer 1 **********************************
        With MapData(x, Y).Graphic(1)
            If (.Started = 1) Then
                If (.SpeedCounter > 0) Then
                    .SpeedCounter = .SpeedCounter - 1
                    If (.SpeedCounter = 0) Then
                        .SpeedCounter = GrhData(.GrhIndex).Speed
                        .FrameCounter = .FrameCounter + 1
                        If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then _
                            .FrameCounter = 1
                    End If
                End If
            End If

            'Figure out what frame to draw (always 1 if not animated)
            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With

        rSourceRect.Left = GrhData(iGrhIndex).sX
        rSourceRect.Top = GrhData(iGrhIndex).sY
        rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight

        'El width fue hardcodeado para speed!
        Call BackBufferSurface.BltFast( _
                ((32 * ScreenX) - 32) + PixelOffsetX, _
                ((32 * ScreenY) - 32) + PixelOffsetY, _
                SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), _
                rSourceRect, _
                DDBLTFAST_WAIT)
        '******************************************
        'Layer 2 **********************************
        If MapData(x, Y).Graphic(2).GrhIndex <> 0 Then
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(x, Y).Graphic(2), _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    1, _
                    1)
        End If
        '******************************************
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y


'busco que nombre dibujar
Call ConvertCPtoTP(frmMain.MainViewShp.Left, frmMain.MainViewShp.Top, frmMain.MouseX, frmMain.MouseY, nX, nY)


'Draw Transparent Layers  (Layer 2, 3)
ScreenY = 8
For Y = minY + 8 To maxY - 1
    ScreenX = 5
    For x = minX + 5 To maxX - 5
        If x > 100 Or x < -3 Then Exit For
        iPPx = 32 * ScreenX - 32 + PixelOffsetX
        iPPy = 32 * ScreenY - 32 + PixelOffsetY

        'Object Layer **********************************
        If MapData(x, Y).ObjGrh.GrhIndex <> 0 Then
'            If Y > UserPos.Y Then
'                Call DDrawTransGrhtoSurfaceAlpha( _
'                        BackBufferSurface, _
'                        MapData(X, Y).ObjGrh, _
'                        iPPx, iPPy, 1, 1)
'            Else
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(x, Y).ObjGrh, _
                        iPPx, iPPy, 1, 1)
'            End If
        End If
        '***********************************************
        'Char layer ************************************
        If MapData(x, Y).charindex <> 0 Then
            TempChar = charlist(MapData(x, Y).charindex)
            PixelOffsetXTemp = PixelOffsetX
            PixelOffsetYTemp = PixelOffsetY

            Moved = 0
            'If needed, move left and right
            If TempChar.MoveOffset.x <> 0 Then
                TempChar.Body.Walk(TempChar.Heading).Started = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.x
               If FeerRLZ = True Then
                    TempChar.MoveOffset.x = TempChar.MoveOffset.x - (2 * Sgn(TempChar.MoveOffset.x))
                Else
                    TempChar.MoveOffset.x = TempChar.MoveOffset.x - (8 * Sgn(TempChar.MoveOffset.x))
                End If
                Moved = 1
            End If
            'If needed, move up and down
            If TempChar.MoveOffset.Y <> 0 Then
                TempChar.Body.Walk(TempChar.Heading).Started = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
               If FeerRLZ = True Then
                    TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (Sgn(TempChar.MoveOffset.Y))
                Else
                    TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (8 * Sgn(TempChar.MoveOffset.Y))
                End If
                Moved = 1
            End If
            'If done moving stop animation
            If Moved = 0 And TempChar.Moving = 1 Then
                TempChar.Moving = 0
                TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                TempChar.Body.Walk(TempChar.Heading).Started = 0
                TempChar.Arma.WeaponWalk(TempChar.Heading).FrameCounter = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                TempChar.Escudo.ShieldWalk(TempChar.Heading).FrameCounter = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 0
            End If
            
'[ANIM ATAK]
            If TempChar.Arma.WeaponAttack > 0 Then
                TempChar.Arma.WeaponAttack = TempChar.Arma.WeaponAttack - 1
                If TempChar.Arma.WeaponAttack = 0 Then
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                End If
            End If
            '[/ANIM ATAK]
            
            '[ANIM ESCUDO]
            If TempChar.Escudo.ShieldAttack > 0 Then
                TempChar.Escudo.ShieldAttack = TempChar.Escudo.ShieldAttack - 1
                If TempChar.Escudo.ShieldAttack = 0 Then
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 0
                End If
            End If
            
            'Dibuja solamente players
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Or (UCase$(TempChar.Nombre) = UCase$(UserName) Or InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") And UserNavegando = True) Then
                If Not charlist(MapData(x, Y).charindex).invisible Then
                
                #If (ConAlfaB = 1) Then
                        If TempChar.Aura.GrhIndex Then
                            Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, TempChar.Aura, _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((34 * ScreenY) - 34) + PixelOffsetYTemp), _
                                    1, 0)
    #End If
End If
               
                            Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
                            Call DDrawTransGrhtoSurface( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.x, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
                
If AuraActivada = True Then

If frmRendimiento.Auris.value = 1 Then
Else
If TempChar.Aura.GrhIndex Then
Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, TempChar.Aura, (((32 * ScreenX) - 32) + PixelOffsetXTemp), (((34 * ScreenY) - 34) + PixelOffsetYTemp), 1, 0)
End If

End If
End If
                
                            Call DDrawTransGrhtoSurface( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.x, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
#If SeguridadAlkon Then
                    If Not MI(CualMI).IsInvisible(MapData(x, Y).charindex) Then
#End If
                        '[CUERPO]'
                            Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
                        '[CABEZA]'
                            Call DDrawTransGrhtoSurface( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.x, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
                        '[Casco]'
                            If TempChar.Casco.Head(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Casco.Head(TempChar.Heading), _
                                        iPPx + TempChar.Body.HeadOffset.x, _
                                        iPPy + TempChar.Body.HeadOffset.Y, _
                                        1, 0)
                            End If
                        '[ARMA]'
                            If TempChar.Arma.WeaponWalk(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Arma.WeaponWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                        '[Escudo]'
                            If TempChar.Escudo.ShieldWalk(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                    
                    
                             If Nombres Then
                                'ya estoy dibujando SOLO si esta visible
                                'If TempChar.invisible = False And Not MI(CualMI).IsInvisible(MapData(X, Y).CharIndex) Then
                                    If TempChar.Nombre <> "" Then
                                        Dim lCenter As Long
                                        'Call Dialogos.DrawText(iPPx - 30, iPPy + 60, "mi:" & IIf(MI(CualMI).IsInvisible(MapData(X, Y).CharIndex), "1", "0") & " .i:" & IIf(TempChar.invisible, "1", "0") & "  X,Y:" & X & "," & Y, RGB(ColoresPJ(5).r, ColoresPJ(5).G, ColoresPJ(5).B))
                                        If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                            lCenter = (frmMain.TextWidth(Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                            Dim sClan As String
                                            sClan = mid(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                            
'Mithrandir sistema de status
Select Case TempChar.priv
Case 0
'Mithrandir
Select Case TempChar.EsStatus
Case 0
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(48).r, ColoresPJ(48).G, ColoresPJ(48).b))
lCenter = (frmMain.TextWidth(sClan) / 2) - 16
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(48).r, ColoresPJ(48).G, ColoresPJ(48).b))
Case 1
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).b))
lCenter = (frmMain.TextWidth(sClan) / 2) - 16
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).b))
Case 2
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).b))
lCenter = (frmMain.TextWidth(sClan) / 2) - 16
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).b))
'Consejo
Case 3
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(5).r, ColoresPJ(5).G, ColoresPJ(5).b))
lCenter = (frmMain.TextWidth(sClan) / 2) - 16
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(5).r, ColoresPJ(5).G, ColoresPJ(5).b))
'Consejo caos
Case 4
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).b))
lCenter = (frmMain.TextWidth(sClan) / 2) - 16
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).b))
End Select
'Mithrandir
Case Is > 0 'admin
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
lCenter = (frmMain.TextWidth(sClan) / 2) - 16
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
Case Else 'el resto
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
lCenter = (frmMain.TextWidth(sClan) / 2) - 16
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
End Select
'Mithrandir sistema de status
                                        Else
lCenter = (frmMain.TextWidth(TempChar.Nombre) / 2) - 16
Select Case TempChar.priv
Case 0
'Mithrandir
Select Case TempChar.EsStatus
Case 0
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(48).r, ColoresPJ(48).G, ColoresPJ(48).b))
Case 1
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).b))
Case 2
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).b))
'Consejo
Case 3
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(5).r, ColoresPJ(5).G, ColoresPJ(5).b))
'Consejo caos
Case 4
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).b))
End Select
'Mithrandir
Case Else
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
End Select
                                        End If
                                    End If
                                'End If  'enidf nI
                             End If
#If SeguridadAlkon Then
                    Else
                        Do While True
                            Call MsgBox("WOAAAAA CHEATER!!! Ahora te deben estar matando de lo lindo ;)" & vbNewLine & "Aprieta OK para salir", vbCritical + vbOKOnly, ":D")
                            Call MsgBox("no, mejor no salimos")
                        Loop
                    End If  'end if not mi.isi
#End If
                End If  'end if ~in

If TempChar.priv <> 0 Then
If TempChar.priv = 2 Then 'SEMI DIOS
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(30, 255, 30))
                                Call Dialogos.DrawText(iPPx - lCenter - 33, iPPy + 43, "<Event Master>", RGB(30, 255, 30))
ElseIf TempChar.priv = 3 Then 'DIOS
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(250, 250, 150))
                                Call Dialogos.DrawText(iPPx - lCenter - 2, iPPy + 43, "<DIOS>", RGB(250, 250, 150))
ElseIf TempChar.priv = 4 Then 'ADMIN
Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(255, 255, 0))
            Call Dialogos.DrawText(iPPx - lCenter - 33, iPPy + 43, "<Administrador>", RGB(255, 255, 0))
End If
End If


                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + TempChar.Body.HeadOffset.x), _
                            (iPPy + TempChar.Body.HeadOffset.Y), _
                            MapData(x, Y).charindex)
                End If
                
                
            Else '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + TempChar.Body.HeadOffset.x), _
                            (iPPy + TempChar.Body.HeadOffset.Y), _
                            MapData(x, Y).charindex)
                End If

                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        TempChar.Body.Walk(TempChar.Heading), _
                        iPPx, iPPy, 1, 1)
            End If '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then


            'Refresh charlist
            charlist(MapData(x, Y).charindex) = TempChar

            'BlitFX (TM)
            If charlist(MapData(x, Y).charindex).Fx <> 0 Then
#If (ConAlfaB = 1) Then
                Call DDrawTransGrhtoSurfaceAlpha( _
                        BackBufferSurface, _
                        FxData(TempChar.Fx).Fx, _
                        iPPx + FxData(TempChar.Fx).OffsetX, _
                        iPPy + FxData(TempChar.Fx).OffsetY, _
                        1, 1, MapData(x, Y).charindex)
#Else
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        FxData(TempChar.Fx).Fx, _
                        iPPx + FxData(TempChar.Fx).OffsetX, _
                        iPPy + FxData(TempChar.Fx).OffsetY, _
                        1, 1, MapData(x, Y).charindex)
#End If
            End If
        End If '<-> If MapData(X, Y).CharIndex <> 0 Then
        '*************************************************
        'Layer 3 *****************************************
        If MapData(x, Y).Graphic(3).GrhIndex <> 0 Then
            'Draw
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(x, Y).Graphic(3), _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    1, 1)
        End If
        If Abs(nX - x) < 1 And (Abs(nY - Y)) < 1 Then
    If MapData(x, Y).ObjGrh.GrhIndex <> 0 Then
    If frmRendimiento.Nomedigas.value = 0 Then
    Else
            Dialogos.DrawText frmMain.MouseX + 200, frmMain.MouseY + 120, MapData(x, Y).ObjName, vbWhite
    End If
    End If
End If
        '************************************************
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y

If frmRendimiento.TechosActivados.value = 1 Then
' TECHOS
If Not bTecho Then
        ScreenY = 5
    For Y = minY + 5 To maxY - 1
        If Y > 0 And Y < 101 Then 'In map bounds
            ScreenX = 5
        For x = minX + 5 To maxX
            If Y > 0 And Y < 101 Then 'In map bounds
    If MapData(x, Y).Graphic(4).GrhIndex <> 0 Then
        'Draw
        Call DDrawTransGrhtoSurface( _
        BackBufferSurface, _
        MapData(x, Y).Graphic(4), _
        ((32 * ScreenX) - 32) + PixelOffsetX, _
        ((32 * ScreenY) - 32) + PixelOffsetY, _
        1, 1)
    End If
End If
        ScreenX = ScreenX + 1
    Next x
End If
        ScreenY = ScreenY + 1
    Next Y
 
Else
 
        ScreenY = 5
            For Y = minY + 5 To maxY - 1
        If Y > 0 And Y < 101 Then 'In map bounds
            ScreenX = 5
            For x = minX + 5 To maxX
        If Y > 0 And Y < 101 Then 'In map bounds
If MapData(x, Y).Graphic(4).GrhIndex <> 0 Then
            'Draw
            Call DDrawTransGrhtoSurfaceAlpha( _
            BackBufferSurface, _
            MapData(x, Y).Graphic(4), _
            ((32 * ScreenX) - 32) + PixelOffsetX, _
            ((32 * ScreenY) - 32) + PixelOffsetY, _
            1, 1)
        End If
    End If
ScreenX = ScreenX + 1
    Next x
End If
    ScreenY = ScreenY + 1
    Next Y
End If
End If

If frmRendimiento.TechosActivados.value = 0 Then
If Not bTecho Then
    'Draw blocked tiles and grid
    ScreenY = 5
    For Y = minY + 5 To maxY - 1
        ScreenX = 5
        For x = minX + 5 To maxX
            'Check to see if in bounds
            If x < 101 And x > 0 And Y < 101 And Y > 0 Then
                If MapData(x, Y).Graphic(4).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(x, Y).Graphic(4), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
                End If
            End If
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next Y
End If
End If





Dim PP As RECT

PP.Left = 0
PP.Top = 0
PP.Right = WindowTileWidth * TilePixelWidth
PP.Bottom = WindowTileHeight * TilePixelHeight


End Sub
 

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2006
'Actualiza todos los sonidos del mapa.
'**************************************************************
    
    DoFogataFx
End Function


Function HayUserAbajo(ByVal x As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And charlist(UserCharIndex).Pos.Y <= Y
        
End If
End Function

Function PixelPos(ByVal x As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    PixelPos = (TilePixelWidth * x) - TilePixelWidth
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
          
    
    'We are done!
    AddtoRichTextBox frmCargando.status, "Hecho.", , , , 1, , False
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"

'Set intial user position
UserPos.x = MinXBorder
UserPos.Y = MinYBorder

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hwnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
    .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If ClientSetup.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck



Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs


AddtoRichTextBox frmCargando.status, "Cargando Gráficos....", 0, 0, 0, , , True
Call LoadGraphics

InitTileEngine = True

End Function

Sub ShowNextFrame()
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer
    
    '****** Set main view rectangle ******
    GetWindowRect DisplayFormhWnd, MainViewRect
    
    With MainViewRect
        .Left = .Left + MainViewLeft
        .Top = .Top + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    If EngineRun Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
         If FeerRLZ = True Then
          OffsetCounterX = (OffsetCounterX - (2 * Sgn(AddtoUserPos.x)))
          Else
          OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.x)))
          End If
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                OffsetCounterX = 0
                AddtoUserPos.x = 0
                UserMoving = 0
            End If
        '****** Move screen Up and Down if needed ******
        ElseIf AddtoUserPos.Y <> 0 Then
         If FeerRLZ = True Then
            OffsetCounterY = OffsetCounterY - (2 * Sgn(AddtoUserPos.Y))
          Else
            OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.Y))
          End If
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = 0
            End If
        End If

        '****** Update screen ******
        Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

  Call Dialogos.DrawText(750, 260, "FPS: " & FramesPerSec, vbYellow)
  
  If CrearonTD = True Then
  Call Dialogos.DrawText(750, 286, "Alpha: " & TDAlpha, vbYellow)
  Call Dialogos.DrawText(750, 300, "Beta: " & TDBeta, vbYellow)
  End If
  
        If frmRendimiento.Check1.value = 1 Then
            EstiloDeNombres = 1
        Else
            EstiloDeNombres = 0
        End If
        
        If IsSeguroC = True Then
        Call Dialogos.DrawText(260, 258, "Seguro de Clan Activado", vbGreen)
        End If
        
        If IsSeguroC = False Then
        Call Dialogos.DrawText(250, 258, "Seguro de Clan Desactivado", vbRed)
        End If
        
        If IsSeguroR = True Then
        Call Dialogos.DrawText(260, 272, "Seguro de Resurreccion Activado", vbGreen)
        End If
        
        If IsSeguroR = False Then
        Call Dialogos.DrawText(260, 272, "Seguro de Resurreccion Desactivado", vbRed)
        End If
        
        If SeguroCvc = True Then
        Call Dialogos.DrawText(260, 286, "Seguro de Cvc Activado", vbGreen)
        Else
        Call Dialogos.DrawText(260, 286, "Seguro de Cvc Desactivado", vbRed)
        End If
        
        Call Dialogos.DrawText(260, 300, "Usuarios Online: " & UsersOns, vbRed)
        Call Dialogos.MostrarTexto
        Call DibujarCartel
        
        Call DialogosClanes.Draw(Dialogos)
        
        
        Call DrawBackBufferSurface
        
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, Index As Integer)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1) 'Anim Invent
Call CrearGrh(534, 2) 'Animacion de teleport
End Sub

Function ControlVelocidad(ByVal lastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - lastTime > 20)
End Function


#If ConAlfaB Then

Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
    Dim rRect As RECT
    
    Surface.GetSurfaceDesc ddsdDest
    
    With rRect
        .Left = 0
        .Top = 0
        .Right = ddsdDest.lWidth
        .Bottom = ddsdDest.lHeight
    End With
    
    If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
        Modo = 0
    ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
        Modo = 1
    Else
        Modo = 2
    End If
    
    Dim DstLock As Boolean
    DstLock = False
    
    On Local Error GoTo HayErrorAlpha
    
    Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
    DstLock = True
    
    Surface.GetLockedArray dArray()
    Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
        ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
        Modo)
    
HayErrorAlpha:
    If DstLock = True Then
        Surface.Unlock rRect
        DstLock = False
    End If
End Sub

#End If
Public Sub DibujarMiniMapa()
On Error Resume Next
Dim map_x, map_y As Byte
    For map_y = 1 To 100
        For map_x = 1 To 100
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
            End If
            If MapData(map_x, map_y).Graphic(2).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(2).GrhIndex).MiniMap_color
            End If
            If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
            End If
        Next map_x
    Next map_y
    frmMain.Minimap.Refresh
End Sub

Public Sub MovemosUserMap()
'%%%%%%%%%%%%%%%%%By Nait%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Comentarios:
'Con esto no llamamos al minimap siempre que damos un paso mucho mas lindo..
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
With frmMain.UserM
.Left = UserPos.x - 3
.Top = UserPos.Y - 3
End With
End Sub
Public Sub MovemosUserMapa()
'%%%%%%%%%%%%%%%%%By Nait%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Comentarios:
'Con esto no llamamos al minimap siempre que damos un paso mucho mas lindo..
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
With frmMap.Shape1
If UserMap = 1 Then
.Top = 2280
.Left = 3360
ElseIf UserMap = 2 Then
.Top = 2280
.Left = 3000
ElseIf UserMap = 3 Then
.Top = 2760
.Left = 3000
ElseIf UserMap = 4 Then
.Top = 2280
.Left = 2250
ElseIf UserMap = 5 Then
.Top = 2280
.Left = 2040
ElseIf UserMap = 6 Then
.Top = 1800
.Left = 2040
ElseIf UserMap = 11 Then
.Top = 2760
.Left = 1680
ElseIf UserMap = 15 Then
.Top = 2760
.Left = 2040
ElseIf UserMap = 16 Then
.Top = 3240
.Left = 2040
ElseIf UserMap = 17 Then
.Top = 3720
.Left = 2040
ElseIf UserMap = 22 Then
.Top = 2280
.Left = 1680
ElseIf UserMap = 31 Then
.Top = 2760
.Left = 1200
ElseIf UserMap = 32 Then
.Top = 2760
.Left = 2520
ElseIf UserMap = 33 Then
.Top = 3600
.Left = 1680
ElseIf UserMap = 34 Then
.Top = 1920
.Left = 1680
ElseIf UserMap = 90 Then
.Top = 3240
.Left = 2520
ElseIf UserMap = 40 Then
.Top = 1440
.Left = 2040
ElseIf UserMap = 107 Then
.Top = 4080
.Left = 2040
ElseIf UserMap = 108 Then
.Top = 4560
.Left = 2040
ElseIf UserMap = 109 Then
.Top = 4560
.Left = 2520
ElseIf UserMap = 111 Then
.Top = 4240
.Left = 2520
Else
.Visible = False
End If
End With
End Sub

Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
On Error Resume Next
    Set PrimarySurface = Nothing
    Set PrimaryClipper = Nothing
    Set BackBufferSurface = Nothing
   
    Set DirectDraw = Nothing
   
    Set DirectX = Nothing
End Sub
