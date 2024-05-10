VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTickCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Dim curFreq     As Currency
Dim curStart    As Currency
Dim curEnd      As Currency
Dim dblResult   As Double

Public Function SetMarker() As Boolean

    On Error GoTo err_handler:
    
    QueryPerformanceFrequency curFreq
    QueryPerformanceCounter curStart
    
    SetMarker = True
    
exit_func:
    Exit Function
    
err_handler:
    Resume exit_func:
    
End Function

Public Function GetElapsedTime(Optional ByVal bRoundTime As Boolean = False) As Double
  
    On Error GoTo err_handler:
    
    GetElapsedTime = -1
    
    If (curFreq = -1) Or (curStart = -1) Then
       
        GoTo exit_func:
    End If
    
    QueryPerformanceCounter curEnd
    dblResult = (curEnd - curStart) / curFreq
    
    If bRoundTime Then
        dblResult = Format(dblResult, "#.##")
    End If
    
    GetElapsedTime = dblResult
    
exit_func:
    Exit Function
    
err_handler:
    Resume exit_func:
    
End Function

Private Sub Class_Initialize()
    curFreq = -1
    curStart = -1
    
    Call SetMarker
End Sub




