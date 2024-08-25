VERSION 5.00
Begin VB.Form cmdClear1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.Activevb.de - Base64"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8865
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdClear2 
      Caption         =   "Clear"
      Height          =   330
      Left            =   8010
      TabIndex        =   7
      Top             =   4530
      Width           =   800
   End
   Begin VB.CommandButton cmdClear1 
      Caption         =   "Clear"
      Height          =   330
      Left            =   2625
      TabIndex        =   6
      Top             =   4530
      Width           =   800
   End
   Begin VB.OptionButton optCodeType 
      Caption         =   "Encryption"
      Height          =   195
      Index           =   1
      Left            =   1305
      TabIndex        =   5
      Top             =   105
      Width           =   1125
   End
   Begin VB.OptionButton optCodeType 
      Caption         =   "Base64"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   90
      Width           =   870
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
      Height          =   3990
      Left            =   3570
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   450
      Width           =   5235
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Decode"
      Height          =   330
      Left            =   3585
      TabIndex        =   2
      Top             =   4530
      Width           =   1485
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Encode"
      Height          =   330
      Left            =   165
      TabIndex        =   1
      Top             =   4530
      Width           =   1485
   End
   Begin VB.TextBox txtDecoded 
      Height          =   4005
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   435
      Width           =   3285
   End
End
Attribute VB_Name = "cmdClear1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

' Autor: K. Langbein Klaus@ActiveVB.de

' Beschreibung: Demonstration der Base64-Kodierung und einer einfachen
' Methode der Verschluesselung mit dem gleichen Algorithmus.
Option Explicit

Dim UsedCode As String

Private Sub chkEncrypt_Click()
    If chkEncrypt.Value = 1 Then
        UsedCode = codeB
        chkBase64.Value = 0
    End If
End Sub

Function TestText$()
    Dim OutText As String
    Dim Oneline As String
    Dim k As Long
    Dim i As Long
    
    Randomize
    For i = 65 To 82
        k = Rnd * 30 + 1
        Oneline = String$(k, i) + vbCrLf
        OutText = OutText + Oneline
    Next i
    
    TestText = OutText
End Function

Private Sub cmdClear1_Click()
    txtDecoded = ""
End Sub

Private Sub cmdClear2_Click()
    txtEncoded = ""
End Sub

Private Sub Command4_Click()
    Dim test As String
    
    test = txtDecoded.text
    test = base64_encode(B64(), test)
    txtEncoded.text = TextBlock(test, 45)
End Sub

Private Sub Command5_Click()
    Dim test As String
    
    test = txtEncoded.text
    test = RemoveCRLF(test)
    test = base64_decode(Rev64, test)
    txtDecoded.text = test
End Sub

Private Sub Form_Load()
    optCodeType(0).Value = -1
    txtDecoded.text = TestText$
End Sub

Private Sub optCodeType_Click(Index As Integer)
    Select Case Index
    Case 0
        UsedCode = Base64
    Case 1
        UsedCode = codeB
    End Select
    
    Call IniCode(UsedCode)
End Sub


