VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.Activevb.de - Base64"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10800
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdClear2 
      Caption         =   "Clear"
      Height          =   330
      Left            =   9480
      TabIndex        =   5
      Top             =   0
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear1 
      Caption         =   "Clear"
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   915
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
      Height          =   4935
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   360
      Width           =   5235
   End
   Begin VB.CommandButton BtnDecode 
      Caption         =   "<-- Decode"
      Height          =   330
      Left            =   5400
      TabIndex        =   2
      Top             =   0
      Width           =   1485
   End
   Begin VB.CommandButton BtnEncode 
      Caption         =   "Encode -->"
      Height          =   330
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   1485
   End
   Begin VB.TextBox txtDecoded 
      Height          =   4965
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   5205
   End
End
Attribute VB_Name = "Form1"
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
Private m_TestText As String

Private Sub Form_Load()
    m_TestText = TestText$
    txtDecoded.text = m_TestText
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
    Dim test As String: test = txtDecoded.text
    txtEncoded.text = TextBlock(Base64_EncodeString(test), 45)
End Sub

Private Sub BtnDecode_Click()
    Dim test As String: test = RemoveCRLF(txtEncoded.text)
    txtDecoded.text = Base64_DecodeString(test)
End Sub

'Private Sub chkEncrypt_Click()
'    If chkEncrypt.Value = 1 Then
'        UsedCode = codeB
'        chkBase64.Value = 0
'    End If
'End Sub
'
'Private Sub optCodeType_Click(Index As Integer)
'    Select Case Index
'    Case 0: UsedCode = Base64
'    Case 1: UsedCode = codeB
'    End Select
'    IniCode UsedCode
'End Sub
'
'
'AAAAAAAAAAAAAAAAAAA
'BBBBBBBBBBBBBBBBBBB
'CCCCCCCCCCCCCCCCCCC
'DDDDDDDDDDDDDDDDD
'EEEEEEEEEEEEEEEEEEE
'FFFFFFFFFFFFFFFFFFFFFF
'GGGGGGGGGGGGGGGGG
'HHHHHHHHHHHHHHHHH
'IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII
'JJJJJJJJJJJJJJJJJJJJJJJJJJJ
'KKKKKKKKKKKKKKKKKKKK
'LLLLLLLLLLLLLLLLLLLLLLL
'MMMMMMMMMMMMMMM
'NNNNNNNNNNNNNNNNN
'OOOOOOOOOOOOOOOOO
'PPPPPPPPPPPPPPPPPPPP
'QQQQQQQQQQQQQQQQQ
'RRRRRRRRRRRRRRRRR
'
'QUFBQUFBQUFBQUFBQUFBQUFBQQ0KQkJCQkJCQkJCQkJCQ
'kJCQkJCQg0KQ0NDQ0NDQ0NDQ0NDQ0NDQ0NDQw0KRERERE
'REREREREREREREREQNCkVFRUVFRUVFRUVFRUVFRUVFRUU
'NCkZGRkZGRkZGRkZGRkZGRkZGRkZGRkYNCkdHR0dHR0dH
'R0dHR0dHR0dHDQpISEhISEhISEhISEhISEhISA0KSUlJS
'UlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSUlJSU
'lJSUlJSUlJDQpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkp
'KSkoNCktLS0tLS0tLS0tLS0tLS0tLS0tLDQpMTExMTExM
'TExMTExMTExMTExMTExMTA0KTU1NTU1NTU1NTU1NTU1ND
'QpOTk5OTk5OTk5OTk5OTk5OTg0KT09PT09PT09PT09PT0
'9PT08NClBQUFBQUFBQUFBQUFBQUFBQUFBQDQpRUVFRUVF
'RUVFRUVFRUVFRUQ0KUlJSUlJSUlJSUlJSUlJSUlI=

'Oliver Meyer
'T2xpdmVyIE1leWVy

'Oliver Meyer 1
'T2xpdmVyIE1leWVyIDE=
'
