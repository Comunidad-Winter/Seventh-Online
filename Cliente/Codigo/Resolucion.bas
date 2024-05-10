Attribute VB_Name = "Resolucion"
Option Explicit
 
Private oldResHeight As Long
Private oldResWidth As Long
Private oldDepth As Integer
Private oldFrequency As Long
 
 
 
Private Declare Function EnumDisplaySettings Lib "User32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "User32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
 
Private Sub IniciarDXobject(DX As DirectX7, DD As DirectDraw7)
On Error Resume Next
 
Set DX = New DirectX7
Set DD = DirectX.DirectDrawCreate("")
End Sub
 
 
Public Sub SetResolution()
    Call IniciarDXobject(DirectX, DirectDraw)
   
Dim lRes As Long
Dim MidevM As typDevMODE
lRes = EnumDisplaySettings(0, 0, MidevM)
 
If MsgBox("¿Desea jugar en pantalla completa?", vbYesNo) = vbYes Then
 
            MidevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            MidevM.dmPelsWidth = 800
            MidevM.dmPelsHeight = 600
            MidevM.dmBitsPerPel = 16
 
      lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
      
      frmMain.WindowState = 2
 
Else
 
frmMain.WindowState = 0
MidevM.dmFields = DM_BITSPERPEL
MidevM.dmBitsPerPel = 16
lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
 
End If
 
End Sub
