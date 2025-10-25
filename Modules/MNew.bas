Attribute VB_Name = "MNew"
Option Explicit

Public Function TestDummy(ByVal NHouses As Byte, ByVal NChildren As Integer, ByVal IsMarried As Boolean, ByVal NCars As Long, ByVal PSofCars As Single, ByVal DistanceToSun As Double, ByVal BirthDay As Date, ByVal Salary As Currency) As TestDummy
    Set TestDummy = New TestDummy: TestDummy.New_ NHouses, NChildren, IsMarried, NCars, PSofCars, DistanceToSun, BirthDay, Salary
End Function


