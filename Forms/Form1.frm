VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SysStrings"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnPadCentered 
      Caption         =   "PadCentered"
      Height          =   375
      Left            =   11520
      TabIndex        =   30
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton BtnResetPadCentered 
      Caption         =   "Reset"
      Height          =   375
      Left            =   11520
      TabIndex        =   32
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   12840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   31
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton BtnTestByteOrderMark 
      Caption         =   "Test ByteOrderMark >>"
      Height          =   375
      Left            =   10920
      TabIndex        =   28
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Some Tests"
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Resizer 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   26
      Top             =   7800
      Width           =   255
   End
   Begin VB.CommandButton BtnReplaceX 
      Caption         =   "Replace "" ."" -> ""."""
      Height          =   375
      Left            =   840
      TabIndex        =   25
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton BtnPadLeftRight 
      Caption         =   "PadLeftRight"
      Height          =   375
      Left            =   7680
      TabIndex        =   22
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   9000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   24
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton BtnResetPadLeftRight 
      Caption         =   "Reset"
      Height          =   375
      Left            =   7680
      TabIndex        =   23
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   19
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton BtnPadRight 
      Caption         =   "PadRight"
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton BtnResetPadRight 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   16
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton BtnPadLeft 
      Caption         =   "PadLeft"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton BtnResetPadLeft 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   12360
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   12
      Top             =   3840
      Width           =   10575
   End
   Begin VB.CommandButton BtnDeleteMultiWS4 
      Caption         =   "DeleteMultiWS"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton BtnRecursiveReplace 
      Caption         =   "RecursiveReplace"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton BtnResetRecursiveReplace 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton BtnDeleteMultiWS3 
      Caption         =   "DeleteMultiWS"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   7
      Top             =   2880
      Width           =   10575
   End
   Begin VB.CommandButton BtnRemoveChars 
      Caption         =   "RemoveChars"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton BtnResetRemoveChars 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton BtnDeleteMultiWS2 
      Caption         =   "DeleteMultiWS"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   960
      Width           =   10575
   End
   Begin VB.CommandButton BtnDeleteCRLF 
      Caption         =   "DeleteCRLF"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton BtnResetDeleteCRLF 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
   Begin VB.CommandButton BtnDeleteMultiWS 
      Caption         =   "DeleteMultiWS"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnResetDeleteMultiWS 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnTestByteOrderMark_Click()
    Form2.Show
End Sub

Private Sub Command1_Click()
    Dim s As String: s = "Dies ist ein Teststring"
    Dim chars() As Integer
    chars = MString.ToCharArray(s, 14, 5)
    MsgBox ChrW(chars(0)) & " " & ChrW(chars(1)) & " " & ChrW(chars(2)) & " " & ChrW(chars(3)) & " " & ChrW(chars(4))
    If MString.StartsWith(s, "Dies") Then MsgBox "Yes, String s starts with ""Dies"""
End Sub

Private Sub Command2_Click()
    'compare PadLeft and PadLeft2
    'interesting are the edges
    'what is if the original string is longer than the given value totalwidth
    'in .net: der urspr�ngliche String wird zur�ckgegeben
    
    Dim s As String: s = "Dies ist ein String"
    
    s = """" & PadLeft(s, 10) & """"
    MsgBox s
    
    's = "Dies ist ein String"
    's = """" & PadLeft2(s, 10) & """"
    'MsgBox s
    
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    BtnResetDeleteMultiWS_Click
    BtnResetDeleteCRLF_Click
    BtnResetRemoveChars_Click
    BtnResetRecursiveReplace_Click
    BtnResetPadLeft_Click
    BtnResetPadRight_Click
    BtnResetPadLeftRight_Click
End Sub

Private Sub Form_Resize()
    Dim l As Single, T As Single, W As Single, H As Single
    Dim m As Single: m = 8 * Screen.TwipsPerPixelX
    
    T = BtnInfo.Top: W = BtnInfo.Width: H = BtnInfo.Height: l = Me.ScaleWidth - m - W
    If W > 0 And H > 0 Then BtnInfo.Move l, T, W, H
    
    l = Text1.Left: T = Text1.Top: H = Text1.Height: W = Me.ScaleWidth - m - W - l
    If W > 0 And H > 0 Then Text1.Move l, T, W, H
    
    T = Text2.Top: W = Me.ScaleWidth - l - m: H = Text2.Height
    If W > 0 And H > 0 Then Text2.Move l, T, W, H
    
    T = Text3.Top: H = Text3.Height
    If W > 0 And H > 0 Then Text3.Move l, T, W, H
    
    T = Text4.Top: H = Text4.Height
    If W > 0 And H > 0 Then Text4.Move l, T, W, H
    
    W = Resizer.Width: H = Resizer.Height: l = Me.ScaleWidth - W: T = Me.ScaleHeight - H:
    If W > 0 And H > 0 Then Resizer.Move l, T, W, H
End Sub

Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

''' v ############################## v ''' Test 1 ''' v ############################## v '''
Private Sub BtnResetDeleteMultiWS_Click()
    Text1.Text = "This string  contains   many    whitespaces     .      With       DeleteMultiWS        only         single          whitespaces           will            remain            .             "
End Sub

Private Sub BtnDeleteMultiWS_Click()
    Text1.Text = MString.DeleteMultiWS(Text1.Text)
End Sub

Private Sub BtnReplaceX_Click()
    Text1.Text = Replace(Text1.Text, " .", ".")
End Sub

''' v ############################## v ''' Test 2 ''' v ############################## v '''
Private Sub BtnResetDeleteCRLF_Click()
    Text2.Text = "This " & vbLf & "string " & vbCr & "contains" & vbCrLf & "many" & vbCrLf & vbCrLf & "Carrier " & vbCr & vbCr & "Return " & vbCrLf & _
                 "and " & vbLf & vbLf & "Line " & vbLf & vbLf & "Feed" & vbCrLf & "with" & vbCrLf & vbCrLf & "DeleteCRLF" & vbCrLf & vbCrLf & "every" & vbCrLf & vbCrLf & _
                 "CR," & vbCrLf & vbCrLf & "LF," & vbCrLf & vbCrLf & "CRLF" & vbCrLf & vbCrLf & "or " & vbLf & vbCr & "LFCR " & vbLf & vbCr & "will " & vbCrLf & vbCrLf & _
                 "be" & vbCrLf & vbCrLf & "replaced" & vbCrLf & vbCrLf & "with" & vbCrLf & vbCrLf & "one" & vbCrLf & vbCrLf & "whitespace."
End Sub

Private Sub BtnDeleteCRLF_Click()
    Text2.Text = MString.DeleteCRLF(Text2.Text)
End Sub

Private Sub BtnDeleteMultiWS2_Click()
    Text2.Text = MString.DeleteMultiWS(Text2.Text)
End Sub

''' v ############################## v ''' Test 3 ''' v ############################## v '''
Private Sub BtnResetRemoveChars_Click()
    Text3.Text = "This \\// string .... contains \././.\ unwanted #### characters ??? ****" & vbCr & vbLf & " every unwanted character will be repalced with a whitespace"
End Sub

Private Sub BtnRemoveChars_Click()
    Text3.Text = MString.RemoveChars(Text3.Text, "\/?.*#" & vbCr & vbLf)
End Sub

Private Sub BtnDeleteMultiWS3_Click()
    Text3.Text = MString.DeleteMultiWS(Text3.Text)
End Sub

''' v ############################## v ''' Test 4 ''' v ############################## v '''
Private Sub BtnResetRecursiveReplace_Click()
    Text4.Text = "This ws string wsws contains ws the ws word ""w_s"" every wswsws occurance ws of ws ""w_s"" wswsws will ws be ws replaced ws with ws a wswswsws whitespace."
End Sub

Private Sub BtnRecursiveReplace_Click()
    Text4.Text = RecursiveReplace(Text4.Text, "ws", " ")
End Sub

Private Sub BtnDeleteMultiWS4_Click()
    Text4.Text = MString.DeleteMultiWS(Text4.Text)
End Sub

''' v ############################## v ''' Test 5 ''' v ############################## v '''
Private Sub BtnResetPadLeft_Click()
    Randomize
    Dim Value As Currency
    ReDim sa(0 To 9) As String
    Dim i As Long
    For i = 0 To UBound(sa)
        Value = Int(Rnd() * 10& ^ (Rnd * 10&))
        sa(i) = CStr(Value)
    Next
    Text5.Text = Join(sa, vbCrLf)
End Sub

Private Sub BtnPadLeft_Click()
    Dim sa() As String: sa = Split(Text5.Text, vbCrLf)
    Dim i As Long, maxlen As Long
    For i = 0 To UBound(sa)
        maxlen = Max(maxlen, Len(sa(i)))
    Next
    For i = 0 To UBound(sa)
        sa(i) = MString.PadLeft(sa(i), maxlen)
    Next
    Text5.Text = Join(sa, vbCrLf)
End Sub

''' v ############################## v ''' Test 6 ''' v ############################## v '''
Private Sub BtnResetPadRight_Click()
    Randomize
    Dim Value As Double
    ReDim sa(0 To 9) As String
    Dim i As Long
    For i = 0 To UBound(sa)
        Value = CLng(Rnd() * 10) / (10& ^ (Rnd() * 10)) 'i&)
        sa(i) = Format(Value, "0." & String(Rnd * 10, "#"))
    Next
    Text6.Text = Join(sa, vbCrLf)
End Sub

Private Sub BtnPadRight_Click()
    Dim sa() As String: sa = Split(Text6.Text, vbCrLf)
    Dim i As Long, maxlen As Long
    For i = 0 To UBound(sa)
        maxlen = Max(maxlen, Len(sa(i)))
    Next
    For i = 0 To UBound(sa)
        sa(i) = MString.PadRight(sa(i), maxlen)
    Next
    Text6.Text = Join(sa, vbCrLf)
End Sub

''' v ############################## v ''' Test 7 ''' v ############################## v '''
Private Sub BtnResetPadLeftRight_Click()
    Randomize
    Dim Value1 As Currency
    Dim Value2 As Double
    ReDim sa(0 To 9) As String
    Dim i As Long
    For i = 0 To UBound(sa)
        Value1 = Int(Rnd() * 10& ^ (Rnd * 10&))
        Value2 = CLng(Rnd() * 10) / (10& ^ (Rnd() * 10))
        sa(i) = CStr(Value1) & Format(Value2, "." & String(Rnd * 10, "#"))
    Next
    Text7.Text = Join(sa, vbCrLf)
End Sub

Private Sub BtnPadLeftRight_Click()
    Dim sa() As String: sa = Split(Text7.Text, vbCrLf)
    Dim ds As String: ds = GetDecimalSeparator
    Dim i As Long, maxlen1 As Long, maxlen2 As Long
    For i = 0 To UBound(sa)
        Dim sx() As String: sx = Split(sa(i), ds)
        maxlen1 = Max(maxlen1, Len(sx(0)))
        maxlen2 = Max(maxlen2, Len(sx(1)))
    Next
    For i = 0 To UBound(sa)
        sx = Split(sa(i), ds)
        sa(i) = MString.PadLeft(sx(0), maxlen1) & ds & MString.PadRight(sx(1), maxlen2)
    Next
    Text7.Text = Join(sa, vbCrLf)
End Sub

''' v ############################## v ''' Test 7 ''' v ############################## v '''
Private Sub BtnResetPadCentered_Click()
    Randomize
    ReDim sa(0 To 9) As String
    Dim i As Long
    For i = 0 To UBound(sa)
        'Value1 = Int(Rnd() * 10& ^ (Rnd * 10&))
        'Value2 = CLng(Rnd() * 10) / (10& ^ (Rnd() * 10))
        sa(i) = CStr(Value1) & Format(Value2, "." & String(Rnd * 10, "#"))
    Next
    Text8.Text = Join(sa, vbCrLf)
End Sub

Private Sub BtnPadCentered_Click()
    '
End Sub

