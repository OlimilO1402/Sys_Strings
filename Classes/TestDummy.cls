VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_nHouses   As Byte     ' you are a very rich person if you have 255 houses
Private m_nChildren As Integer  ' a person can only have 32767 children so 128 children per house
Private m_IsMarried As Boolean  ' if a person has so many children, it must be married, isn't it
Private m_nCars     As Long     ' a person can never have enough cars but max 2147483647 we could deal with
Private m_PSofCars  As Single   ' the sum of all PS all the persons cars have together
Private m_DistToSun As Double   ' distance from the persons planet to the sun
Private m_BirthD    As Date     ' the persons birthday
Private m_Salary    As Currency ' the amount of salary per hour

Private Sub Class_Initialize()
    m_nHouses = 123
    m_nChildren = 1234
    m_IsMarried = True
    m_nCars = 12345
    m_PSofCars = 123456.7
    m_DistToSun = 12345678901.2346
    m_BirthD = DateSerial(1996, 2, 29)
    m_Salary = 12345.6789@
End Sub

Friend Sub New_(NHouses As Byte, NChildren As Integer, IsMarried As Boolean, NCars As Long, PSofCars As Single, DistanceToSun As Double, BirthDay As Date, Salary As Currency)
    m_nHouses = NHouses: m_nChildren = NChildren: m_IsMarried = IsMarried: m_nCars = NCars: m_PSofCars = PSofCars: m_DistToSun = DistanceToSun: m_BirthD = BirthDay: m_Salary = Salary
End Sub

Public Function Clone() As TestDummy
    Set Clone = New TestDummy: Clone.NewC Me
End Function

Public Sub NewC(other As TestDummy)
    With other
        m_nHouses = .NHouses
        m_nChildren = .NChildren
        m_IsMarried = .IsMarried
        m_nCars = .NCars
        m_PSofCars = .PSofCars
        m_DistToSun = .DistanceToSun
        m_BirthD = .BirthDay
        m_Salary = .Salary
    End With
End Sub

Public Sub SetParams(Optional NHouses, Optional NChildren, Optional IsMarried, Optional NCars, Optional PSofCars, Optional DistanceToSun, Optional BirthDay, Optional Salary)
    If Not IsMissing(NHouses) Then m_nHouses = NHouses
    If Not IsMissing(NChildren) Then m_nChildren = NChildren
    If Not IsMissing(IsMarried) Then m_IsMarried = IsMarried
    If Not IsMissing(NCars) Then m_nCars = NCars
    If Not IsMissing(PSofCars) Then m_PSofCars = PSofCars
    If Not IsMissing(DistanceToSun) Then m_DistToSun = DistanceToSun
    If Not IsMissing(BirthDay) Then m_BirthD = BirthDay
    If Not IsMissing(Salary) Then m_Salary = Salary
End Sub

Public Property Get NHouses() As Byte
    NHouses = m_nHouses
End Property

Public Property Get NChildren() As Integer
    NChildren = m_nChildren
End Property

Public Property Get IsMarried() As Boolean
    IsMarried = m_IsMarried
End Property

Public Property Get NCars() As Long
    NCars = m_nCars
End Property

Public Property Get PSofCars() As Single
    PSofCars = m_PSofCars
End Property

Public Property Get DistanceToSun() As Double
    DistanceToSun = m_DistToSun
End Property

Public Property Get BirthDay() As Date
    BirthDay = m_BirthD
End Property

Public Property Get Salary() As Currency
    Salary = m_Salary
End Property

Public Function ToStr() As String
    Dim s As String
    s = s & "nHouses   : " & m_nHouses & vbCrLf
    s = s & "nChildren : " & m_nChildren & vbCrLf
    s = s & "IsMarried : " & m_IsMarried & vbCrLf
    s = s & "nCars     : " & m_nCars & vbCrLf
    s = s & "PSofCars  : " & m_PSofCars & vbCrLf
    s = s & "DistToSun : " & m_DistToSun & vbCrLf
    s = s & "Birthday  : " & m_BirthD & vbCrLf
    s = s & "Salary    : " & m_Salary & vbCrLf
    ToStr = s
End Function
