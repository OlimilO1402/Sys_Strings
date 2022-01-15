VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Test_ConvertFromAnsiString
    'Test_ConvertFromWideString
    
End Sub

Sub Test_ConvertFromAnsiString()
    
    Dim sSrc As String: sSrc = StrConv("Dies ist ein String.", vbFromUnicode)
    
    Dim sDst As String: sDst = Module2.StringFromLPCStr(StrPtr(sSrc))
    
    MsgBox LenB(sDst) & " " & sDst
    
End Sub

Sub Test_ConvertFromWideString()
    
    Dim sSrc As String: sSrc = "Dies ist ein String."
    
    Dim sDst As String: sDst = Module2.StringFromLPCWStr(StrPtr(sSrc))
    
    MsgBox LenB(sDst) & " " & sDst
    
End Sub


