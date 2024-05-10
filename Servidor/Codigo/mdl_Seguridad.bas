Attribute VB_Name = "mdl_Seguridad"
Option Explicit

Function Ofdjclsdkf(ByVal s As String) As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
Dim p As String
r = ""
p = "PzH642!5hQ19!"
If Len(p) > 0 Then
For i = 1 To Len(s)
C1 = Asc(mid(s, i, 1))
If i > Len(p) Then
C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
Else
C2 = Asc(mid(p, i, 1))
End If
C1 = C1 + C2 + 64
If C1 > 255 Then C1 = C1 - 256
r = r + Chr(C1)
Next i
Else
r = s
End If
Ofdjclsdkf = r
End Function

Function Lfdjcnzsmfdd(ByVal s As String) As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
Dim p As String
p = "PzH642!5hQ19!"
r = ""
If Len(p) > 0 Then
For i = 1 To Len(s)
C1 = Asc(mid(s, i, 1))
If i > Len(p) Then
C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
Else
C2 = Asc(mid(p, i, 1))
End If
C1 = C1 - C2 - 64
If Sgn(C1) = -1 Then C1 = 256 + C1
r = r + Chr(C1)
Next i
Else
r = s
End If
Lfdjcnzsmfdd = r
End Function


