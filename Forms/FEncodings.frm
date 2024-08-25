VERSION 5.00
Begin VB.Form FEncodings 
   Caption         =   "Encodings"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   7095
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton OptJavaScr 
      Caption         =   "JavaScript \u...."
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton OptBase64 
      Caption         =   "Base64"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtEncoded 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   840
      Width           =   5175
   End
   Begin VB.CommandButton BtnDecode 
      Caption         =   "<-- Decode"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtDecoded 
      Height          =   5295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Text            =   "FEncodings.frx":0000
      Top             =   840
      Width           =   5190
   End
   Begin VB.CommandButton BtnEncode 
      Caption         =   "Encode -->"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FEncodings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private m_TestText As String

Private Sub Form_Load()
    OptBase64.Value = True
End Sub

Private Sub OptBase64_Click()
    txtDecoded.text = vbNullString
    txtDecoded.text = TestText$
End Sub

Private Sub OptJavaScr_Click()
    txtDecoded.text = vbNullString
    txtEncoded.text = "\u00c4\u00d6\u00dc\u00e4\u00f6\u00fc\u00df, " & vbCrLf & _
                      "\u00c4hren, \u00d6ltanker, \u00dcberschrift, " & vbCrLf & _
                      "F\u00e4rberkamille und Wilde M\u00f6hre " & vbCrLf & _
                      "\u00fcbernehmen die Hauptstra\u00dfe\u\u\u"
    '    MsgBox JSEscaped_Decode("")
'    MsgBox JSEscaped_Decode("\u")
'    MsgBox JSEscaped_Decode("\u00200")
'    MsgBox JSEscaped_Decode("\u00c4\u00d6\u00dc\u00e4\u00f6\u00fc\u00df, \u00c4hren, \u00d6ltanker, \u00dcberschrift, F\u00e4rberkamille und Wilde M\u00f6hre \u00fcbernehmen die Hauptstra\u00dfe\u\u\u")

End Sub

Function TestText$()
    Dim s As String
    Dim i As Long, k As Long
    Randomize
    For i = 65 To 82
        k = Rnd * 30 + 1
        s = s & String$(k, i) + vbCrLf
    Next i
    TestText = s
End Function

Private Sub cmdClear1_Click()
    txtDecoded = ""
End Sub

Private Sub cmdClear2_Click()
    txtEncoded = ""
End Sub

Private Sub BtnEncode_Click()
    If OptBase64.Value Then
        txtEncoded.text = TextBlock(Base64_EncodeString(txtDecoded.text), 45)
    Else
        MsgBox "not yet implemented"
    End If
End Sub

Private Sub BtnDecode_Click()
    If OptBase64.Value Then
        txtDecoded.text = MString.Base64_DecodeString(txtEncoded.text)
    Else
        txtDecoded.text = MString.JSEscaped_Decode(txtEncoded.text)
    End If
End Sub


Public Function TextBlock(text As String, ByVal nChars As Long) As String
    ' Erzeugung eines Textblockes mit konstanter
    ' Zeilenlaenge fuer die Darstellung. Dies wird bei
    ' Mailattachments auch gemacht.

    Dim s As String
    Dim Oneline As String
    Dim i As Long
    
    For i = 1 To Len(text) Step nChars
        Oneline = Mid$(text, i, nChars) & vbCrLf
        s = s + Oneline
    Next i
    
    TextBlock = s
End Function


Public Function RemoveCRLF(text As String) As String
    
    Dim s As String
    Dim Oneline As String
    
    ' Carriage-Return und Line-Feed koennen per Definition
    ' nicht in einem mit Base64 kodierten Text enthalten sein.
    ' Sie werden aber meist nach je 45-60 Zeichen eingefuegt,
    ' um den Text lesbar zu machen. Hier werden sie wieder entfernt.
    
    Dim pos1 As Long: pos1 = 1
    Dim pos2 As Long
    Do
        
        pos2 = InStr(pos1, text, vbCrLf, 1)
        If pos2 > 0 Then
            Oneline = Mid$(text, pos1, pos2 - pos1)
            s = s & Oneline
            pos1 = pos2 + 2
        Else
            Oneline = Mid$(text, pos1)
            s = s + Oneline
        End If
    
    Loop Until pos2 = 0
    
    RemoveCRLF = s
End Function
