VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4395
   ClientLeft      =   19155
   ClientTop       =   3090
   ClientWidth     =   10425
   LinkTopic       =   "Form2"
   ScaleHeight     =   4395
   ScaleWidth      =   10425
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
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Width           =   4575
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
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   4575
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
    Text1.Text = FileReadAllString(FNm)
End Sub

Private Sub Command2_Click()
    Dim FNm As String:  FNm = App.Path & "\StringUTF8.txt"
    Text2.Text = FileReadAllString(FNm)
End Sub

Private Sub Command3_Click()
    Dim FNm As String:  FNm = App.Path & "\StringUTF8_bom.txt"
    Text3.Text = FileReadAllString(FNm)
End Sub

Private Sub Command4_Click()
    Dim FNm As String:  FNm = App.Path & "\StringUTF16BE_bom.txt"
    Text4.Text = FileReadAllString(FNm)
End Sub

Private Sub Command5_Click()
    Dim FNm As String:  FNm = App.Path & "\StringUTF16LE_bom.txt"
    Text5.Text = FileReadAllString(FNm)
End Sub

Function FileReadAllString(FNm As String) As String
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    'Open FNm For Input As FNr
    Open FNm For Binary As FNr
    Dim FileContent As String: FileContent = Space(LOF(FNr))
    Get FNr, , FileContent
    FileReadAllString = FileContent
    GoTo Finally
Catch:
    MsgBox Err.Description
Finally:
    Close FNr
End Function

