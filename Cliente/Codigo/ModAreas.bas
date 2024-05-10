Attribute VB_Name = "ModAreas"

Option Explicit

Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer

Public Sub CambioDeArea(ByVal x As Byte, ByVal y As Byte)
    Dim loopX As Long, loopY As Long
    Dim tempint As Integer
    
    MinLimiteX = (x \ 9 - 1) * 9
    MaxLimiteX = MinLimiteX + 26
    
    MinLimiteY = (y \ 9 - 1) * 9
    MaxLimiteY = MinLimiteY + 26
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                
                If MapData(loopX, loopY).charindex > 0 Then
                    If MapData(loopX, loopY).charindex <> UserCharIndex Then
                        tempint = MapData(loopX, loopY).charindex
                        Call EraseChar(MapData(loopX, loopY).charindex)
                        charlist(tempint).Nombre = loopX & "-" & loopY
                    End If
                End If
                
                'Erase OBJs
                MapData(loopX, loopY).ObjGrh.GrhIndex = 0
            End If
        Next
    Next
    
    Call RefreshAllChars
End Sub

Public Sub ClearMap()
Dim loopX As Long, loopY As Long
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            'Erase NPCs
            If MapData(loopX, loopY).charindex > 0 Then
                If MapData(loopX, loopY).charindex <> UserCharIndex Then
                    Call EraseChar(MapData(loopX, loopY).charindex)
                End If
            End If
            
            'Erase OBJs
            MapData(loopX, loopY).ObjGrh.GrhIndex = 0
        Next
    Next
    
    Call RefreshAllChars
End Sub
