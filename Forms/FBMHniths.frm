VERSION 5.00
Begin VB.Form FBMHniths 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test of searching for the needle in the haystack with an algorithm by Boyer, Moore and Horspool (W)"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   600
      Width           =   9975
   End
End
Attribute VB_Name = "FBMHniths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Haystack As String
Private m_Needle   As String
Private m_Start    As Long

Private Sub Form_Load()
    m_Haystack = "ui ui ui ui ui the needle is in the haystack, go and find the needle"
    m_Needle = "needle"
End Sub

Private Sub BtnReset_Click()
    m_Start = 0
End Sub

Public Sub Debug_Print(ByVal s As String)
    Me.Text1 = Me.Text1 & s & vbCrLf
End Sub

Private Sub Command1_Click()
    Debug_Print m_Haystack
    Debug_Print m_Needle
    If m_Start = 0 Then
        m_Start = MString.FindStr(m_Haystack, m_Needle)
    Else
        m_Start = MString.FindNext
    End If
    Debug_Print "pos: " & m_Start
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move L, t, W, H
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MString.BMH_Clear
End Sub

