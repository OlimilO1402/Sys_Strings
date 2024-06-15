VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4395
   ClientLeft      =   19155
   ClientTop       =   3090
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   ScaleHeight     =   4395
   ScaleWidth      =   11400
   Begin VB.CommandButton Command1 
      Caption         =   "StringAnsiWin1252"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   9135
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   9135
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Width           =   9135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "StringUTF16LE_bom"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "StringUTF16BE_bom"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   9135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   9135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "StringUTF8_bom"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "StringUTF8"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim FNm As String:  FNm = App.Path & "\StringAnsiWindows1252.txt"
    Dim s As String: s = FileReadAllString(FNm)
    Dim bom As EByteOrderMark: bom = IsBOM(s, s)
    If bom = bom_None Then
        s = StrConv(s, vbUnicode)
    End If
'    If bom = bom_UTF_8 Then
'        Dim buffer() As Byte: buffer = s
'        s = MString.ConvertFromUTF8(buffer)
'    End If
    Text1.Text = s
End Sub

Private Sub Command2_Click()
    Dim FNm As String:  FNm = App.Path & "\StringUTF8.txt"
    Dim s As String: s = FileReadAllString(FNm)
    Dim bom As EByteOrderMark: bom = IsBOM(s, s)
    If bom = bom_UTF_8 Then
        Dim buffer() As Byte: buffer = s
        s = MString.ConvertFromUTF8(buffer)
    End If
    Text2.Text = s
End Sub

Private Sub Command3_Click()
    Dim FNm As String:  FNm = App.Path & "\StringUTF8_bom.txt"
    Dim s As String: s = FileReadAllString(FNm)
    Text3.Text = s
End Sub

Private Sub Command4_Click()
    Dim FNm As String:  FNm = App.Path & "\StringUTF16BE_bom.txt"
    Text4.Text = FileReadAllString(FNm)
End Sub

Private Sub Command5_Click()
    Dim FNm As String:  FNm = App.Path & "\StringUTF16LE_bom.txt"
    Dim s As String: s = FileReadAllString(FNm)
    Text5.Text = s
End Sub

Function FileReadAllString(FNm As String) As String
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open FNm For Binary As FNr
    'Dim bom As EByteOrderMark 'Long
    'Dim ibom As Integer
    'Get FNr, , ibom
    'bom = ibom
    Dim u As Long: u = LOF(FNr) - 1
    'If bom <> bom_None Then u = u - MString.E
    ReDim filecontent(0 To u) As Byte
    Get FNr, , filecontent
    Dim s As String
    
    'Select Case bom
    'Case EByteOrderMark.bom_UTF_16_BE
    '    'swap the order around
    '    s = filecontent
    '    MPtr.String_Rotate2 s
    'Case EByteOrderMark.bom_UTF_16_LE
    '    'do nothing, string is perfekt as it should be
    'Case EByteOrderMark.bom_UTF_32_BE
    '    s = StrConv(filecontent, vbFromUnicode)
    '    MPtr.String_Rotate2 s
    'Case EByteOrderMark.bom_UTF_32_LE
    '    s = StrConv(filecontent, vbFromUnicode)
    'Case EByteOrderMark.bom_UTF_7
    '
    'Case EByteOrderMark.bom_UTF_8
    '    s = MString.ConvertFromUTF8(filecontent)
    'Case EByteOrderMark.bom_UTF_EBCDIC
        'sorry no solution yet!
    'Case Else
        s = filecontent
    'End Select
    FileReadAllString = s
    GoTo Finally
Catch:
    MsgBox Err.Description
Finally:
    Close FNr
End Function

