VERSION 5.00
Begin VB.Form FTestTryParseValidate 
   Caption         =   "FTestTryParseValidate"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Salary per hour:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1395
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Day of birth:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Distance to sun:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "PS of all cars:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Is married:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Number of cars:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Number of children:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of houses:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "FTestTryParseValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Result As VbMsgBoxResult
Private m_Person As TestDummy

Private Sub Form_Load()
    m_Result = VbMsgBoxResult.vbCancel
End Sub

Public Function ShowDialog(Person As TestDummy, Owner As Form) As VbMsgBoxResult
    Set m_Person = Person.Clone
    UpdateView m_Person
    Me.Show vbModal, Owner
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    Person.NewC m_Person
End Function

Sub UpdateView(Object)
    Dim VIP As TestDummy: Set VIP = Object
    With VIP
        Text1.text = .NHouses
        Text2.text = .NChildren
        Text3.text = .IsMarried
        Text4.text = .NCars
        Text5.text = .PSofCars
        Text6.text = .DistanceToSun
        Text7.text = .BirthDay
        Text8.text = .Salary
    End With
End Sub

Function UpdateData(Object) As Boolean
    Dim VIP As TestDummy: Set VIP = Object
    Dim bIsOK As Boolean
    With VIP
        Dim bv As Byte:     bv = .NHouses:       Text1.text = MString.Byte_TryParseValidate(Text1.text, Label1.Caption, "", bIsOK, bv):            If Not bIsOK Then Exit Function
        Dim bi As Integer:  bi = .NChildren:     Text2.text = MString.Integer_TryParseValidate(Text2.text, Label2.Caption, "", bIsOK, bi):         If Not bIsOK Then Exit Function
        Dim bb As Boolean:  bb = .IsMarried:     Text3.text = MString.Boolean_TryParseValidate(Text3.text, Label3.Caption, "", bIsOK, bb):         If Not bIsOK Then Exit Function
        Dim bl As Long:     bl = .NCars:         Text4.text = MString.Long_TryParseValidate(Text4.text, Label4.Caption, "", bIsOK, bl):            If Not bIsOK Then Exit Function
        Dim bs As Single:   bs = .PSofCars:      Text5.text = MString.Single_TryParseValidate(Text5.text, Label5.Caption, "0.0000", bIsOK, bs):    If Not bIsOK Then Exit Function
        Dim bd As Double:   bd = .DistanceToSun: Text6.text = MString.Double_TryParseValidate(Text6.text, Label6.Caption, "0.0000000", bIsOK, bd): If Not bIsOK Then Exit Function
        Dim dt As Date:     dt = .BirthDay:      Text7.text = MString.Date_TryParseValidate(Text7.text, Label7.Caption, "", bIsOK, dt):            If Not bIsOK Then Exit Function
        Dim bc As Currency: bc = .Salary:        Text8.text = MString.Currency_TryParseValidate(Text8.text, Label8.Caption, "0.0000", bIsOK, bc):  If Not bIsOK Then Exit Function
    End With
    VIP.New_ bv, bi, bb, bl, bs, bd, dt, bc
    UpdateData = True
End Function

Private Sub Text1_Validate(Cancel As Boolean)
    Dim byt   As Byte:           byt = m_Person.NHouses
    Dim bIsOK As Boolean: Text1.text = MString.Byte_TryParseValidate(Text1.text, Label1.Caption, "", bIsOK, byt)
    If bIsOK Then m_Person.SetParams NHouses:=byt
    Cancel = Not bIsOK
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    Dim iint  As Integer:       iint = m_Person.NChildren
    Dim bIsOK As Boolean: Text2.text = MString.Integer_TryParseValidate(Text2.text, Label2.Caption, "", bIsOK, iint)
    If bIsOK Then m_Person.SetParams NChildren:=iint
    Cancel = Not bIsOK
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    Dim bln   As Boolean:        bln = m_Person.IsMarried
    Dim bIsOK As Boolean: Text3.text = MString.Boolean_TryParseValidate(Text3.text, Label3.Caption, "", bIsOK, bln)
    If bIsOK Then m_Person.SetParams IsMarried:=bln
    Cancel = Not bIsOK
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
    Dim lng   As Long:           lng = m_Person.NCars
    Dim bIsOK As Boolean: Text4.text = MString.Long_TryParseValidate(Text4.text, Label4.Caption, "", bIsOK, lng)
    If bIsOK Then m_Person.SetParams NCars:=lng
    Cancel = Not bIsOK
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
    Dim sng   As Single:         sng = m_Person.PSofCars
    Dim bIsOK As Boolean: Text5.text = MString.Single_TryParseValidate(Text5.text, Label5.Caption, "0.0", bIsOK, sng)
    If bIsOK Then m_Person.SetParams PSofCars:=sng
    Cancel = Not bIsOK
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
    Dim dbl   As Double:         dbl = m_Person.DistanceToSun
    Dim bIsOK As Boolean: Text6.text = MString.Double_TryParseValidate(Text6.text, Label6.Caption, "0.0000000", bIsOK, dbl)
    If bIsOK Then m_Person.SetParams DistanceToSun:=dbl
    Cancel = Not bIsOK
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
    Dim dat   As Date:           dat = m_Person.BirthDay
    Dim bIsOK As Boolean: Text7.text = MString.Date_TryParseValidate(Text7.text, Label7.Caption, "", bIsOK, dat)
    If bIsOK Then m_Person.SetParams BirthDay:=dat
    Cancel = Not bIsOK
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
    Dim cur   As Currency:       cur = m_Person.Salary
    Dim bIsOK As Boolean: Text8.text = MString.Currency_TryParseValidate(Text8.text, Label8.Caption, "0.0000", bIsOK, cur)
    If bIsOK Then m_Person.SetParams Salary:=cur
    Cancel = Not bIsOK
End Sub

Private Sub BtnOK_Click()
    If Not UpdateData(m_Person) Then Exit Sub
    m_Result = vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

