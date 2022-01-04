Attribute VB_Name = "MString"
Option Explicit 'Zeilen: 129

'String-Routinen
Public Function DeleteMultiWS(s As String) As String
    'Replace Recursive Delete Multi WhiteSpace WS
    DeleteMultiWS = Trim$(s)
    If InStr(1, s, "  ") = 0 Then Exit Function
    DeleteMultiWS = Replace(s, "  ", " ")
    DeleteMultiWS = DeleteMultiWS(DeleteMultiWS)
End Function

Public Function DeleteCRLF(s As String) As String
    DeleteCRLF = Trim$(s)
    If InStr(1, s, vbLf) = 0 Then Exit Function
    If InStr(1, s, vbCr) = 0 Then Exit Function
    DeleteCRLF = Replace(Replace(Replace(s, vbCrLf, " "), vbLf, " "), vbCr, " ")
    DeleteCRLF = DeleteCRLF(DeleteCRLF)
End Function

Public Function RemoveChars(ByVal this As String, CharsToRemove As String) As String
    Dim c As String
    Dim i As Long
    RemoveChars = this
    For i = 1 To Len(CharsToRemove)
        c = Mid$(CharsToRemove, i, 1)
        If InStr(1, this, c) Then
            RemoveChars = Replace(RemoveChars, c, vbNullString)
        End If
    Next
End Function

Public Function RecursiveReplace(ByVal Expression As String, ByVal Find As String, ByVal Replace As String) As String
    'Returns a string where all occurances of "Find" in "Expression" are replaced by "Replace".
    'RecursivReplace removes multi Backslashes at once e.g. to replace „\\“ by „\“
    'a normal Replace("C:\\\test\\\path\\\dir\\\file.txt", "\\", "\") returns „C:\\test\\path\\dir\\file.txt“
    ' RecursivReplace("C:\\\test\\\path\\\dir\\\file.txt", "\\", "\") returns „C:\test\path\dir\file.txt“
    
    Dim pos As Long: pos = InStr(1, Expression, Find)
    If pos Then
        Dim r As String: r = VBA.Replace(Expression, Find, Replace)
        'check for stack overflow:
        If (r = Expression) Or (Len(Expression) < Len(r)) Then RecursiveReplace = r: Exit Function
        RecursiveReplace = RecursiveReplace(r, Find, Replace)
    Else
        RecursiveReplace = Expression
    End If
End Function

Public Function RecursiveReplaceSL(ByVal Expression As String, ByVal Find As String, ByVal Replace As String, Optional ByVal Start As Long = 1, Optional ByVal length As Long = -1) As String
    'Uses RecursiveReplace to replace "Find" by "Replace" in a part of "Expression" that starts with "Start" with the length of "Length"
    'check input parameters return early if necessary
    If length < 0 And Start = 1 Then RecursiveReplaceSL = RecursiveReplace(Expression, Find, Replace): Exit Function
    Dim le As Long: le = Len(Expression)
    If Start < 1 Or le < Start Then Exit Function 'return nothing
    If length < 1 Or le < Start + length Then length = le - Start + 1
    
    Dim sl As String: sl = Left$(Expression, Start - 1)
    Dim sm As String: sm = Mid$(Expression, Start, length)
    Dim sr As String: sr = Mid$(Expression, Start + length)
    sm = RecursiveReplace(sm, Find, Replace)
    RecursiveReplaceSL = sl & sm & sr
    'same but shorter and less noise:
    'RecursiveReplaceSL = Left$(Expression, Start - 1) & RecursiveReplace(Mid$(Expression, Start, Length)) & Mid$(Expression, Start, Length)
End Function


'Converters to or from String
'Bool
Public Function BoolToYesNo(ByVal b As Boolean) As String
    BoolToYesNo = IIf(b, " Ja ", "Nein")
End Function

Public Function Double_TryParse(ByVal Value As String, ByRef d_out As Double) As Boolean
Try: On Error GoTo Catch
    Value = Replace(Value, ",", ".")
    d_out = Val(Value)
    Double_TryParse = True
Catch:
End Function

Public Function Single_TryParse(ByVal Value As String, ByRef s_out As Single) As Boolean
Try: On Error GoTo Catch
    Value = Replace(Value, ",", ".")
    s_out = CSng(Val(Value))
    Single_TryParse = True
    Exit Function
Catch:
End Function

Public Function Hex2(ByVal Value As Byte) As String
    Hex2 = Hex(Value): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function

Public Function Hex4(ByVal Value As Integer) As String
    Hex4 = Hex(Value): If Len(Hex4) < 4 Then Hex4 = String(4 - Len(Hex4), "0") & Hex4
End Function

Public Function Hex8(ByVal Value As Long) As String
    Hex8 = Hex(Value): If Len(Hex8) < 8 Then Hex8 = String(8 - Len(Hex8), "0") & Hex8
End Function

Public Function Contains(s As String, ByVal Value As String) As Boolean
    Contains = InStr(1, s, Value) > 0
End Function

Public Function EndsWith(s As String, ByVal Value As String) As Boolean
    EndsWith = Value = Left$(s, Len(Value))
End Function

'?IndexOf("Dies ist ein String", "ein") = 9
'?IndexOf("Dies ist ein String", "ein", 0) = 9
'?IndexOf("Dies ist ein String", "ein", 1) = 9
'?IndexOf("Dies ist ein String", "ein", 2) = 9
'?IndexOf("Dies ist ein String", "en", 2) = -1
Public Function IndexOf(s As String, ByVal Value As String, Optional ByVal startIndex As Long = 0, Optional ByVal Count As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
'Gibt den Null-basierten Index des ersten Vorkommens der angegebenen Zeichenfolge in dieser Instanz an.
'Die Suche beginnt an einer angegebenen Zeichenpopsition.
'Rückgabewerte:
'Die nullbasierte Indexposition von Value vom Anfang der aktuellen Instanz, wenn diese Zeichenfolge gefunden
'wurde, oder -1 wenn sie nicht gefunden wurde. Wenn value leer ist, wird startindex zurückgegeben.
    If startIndex < 0 Then startIndex = 0
    If Len(s) < startIndex Then startIndex = Len(s)
    If Count < 0 Then Count = Len(s) - startIndex
    If Len(s) < startIndex + Count - 1 Then Count = Len(s) - startIndex
    Dim v As String: v = MidB(s, startIndex + 1, (Count + 1) * 2)
    IndexOf = InStr(1, v, Value, Compare) - 1
    If IndexOf > 0 Then IndexOf = startIndex + IndexOf - 1
End Function

Public Function Insert(s As String, ByVal startIndex As Long, ByVal Value As String) As String
    '
End Function

Public Function LastIndexOf(s As String, Value As String, ByVal startIndex As Long, ByVal Count As Long, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
    '
End Function

Public Function GetDecimalSeparator() As String
    Dim d As Double, s As String
    s = Format(d, "0.0")
    GetDecimalSeparator = Mid(s, 2, 1)
End Function

Function PadLeft(this As String, _
                 ByVal totalWidth As Long, _
                 Optional ByVal paddingChar As String) As String
    
    ' Returns the String StrVal with the specified length.
    ' totalWidth: the length of the returned string
    '             if totalWidth is smaller then the length of StrVal then
    '             StrVal will be returned
    ' padChar:    on the left hand side it will be filled up with padChar
    '             if padChar is not specified, the returned string will be
    '             filled up with spaces.
    
    'ist der String länger als totalwidth, wird nur der String zurückgegeben
    'ansonsten wird der String mit der angegebenen Länge zurückgegeben, der
    'String wird nach rechts gerückt, und links mit PadChar aufgefüllt
    'ist PadChar nicht angegeben, so wird mit RSet der String in
    'Spaces eingefügt.
    
    If LenB(paddingChar) Then
        If Len(this) < totalWidth Then
            PadLeft = String$(totalWidth, paddingChar)
            MidB$(PadLeft, totalWidth * 2 - LenB(this) + 1) = this
        Else
            PadLeft = this
        End If
    Else
        PadLeft = Space$(totalWidth)
        RSet PadLeft = this
    End If
End Function

Function PadRight(this As String, _
                  ByVal totalWidth As Long, _
                  Optional ByVal paddingChar As String) As String
    
    ' Returns the String StrVal with the specified length.
    ' totalWidth: the length of the returned string
    '             if totalWidth is smaller then the length of StrVal then
    '             StrVal will be returned
    ' padChar:    on the right hand side it will be filed up with padChar
    '             if padChar is not specified, the returned string will be
    '             filled up with spaces.
    
    'ist der String länger als totalwidth, wird nur der String zurückgegeben
    'ansonsten wird der String mit der angegebenen Länge zurückgegeben, der
    'String wird nach links gerückt, und rechts mit PadChar aufgefüllt
    'ist PadChar nicht angegeben, so wird mit LSet der String in
    'Spaces eingefügt.
    
    If LenB(paddingChar) Then
        If Len(this) < totalWidth Then
            PadRight = String$(totalWidth, paddingChar)
            MidB$(PadRight, 1) = this
        Else
            PadRight = this
        End If
    Else
        PadRight = Space$(totalWidth)
        LSet PadRight = this
    End If
End Function

'Public Function PadLeft(StrVal As String, _
'                        ByVal totalWidth As Long, _
'                        Optional ByVal padChar As String) As String
'
'    ' Returns the String StrVal with the specified length.
'    ' totalWidth: the length of the returned string
'    '             if totalWidth is smaller then the length of StrVal then
'    '             StrVal will be returned
'    ' padChar:    on the left hand side it will be filled up with padChar
'    '             if padChar is not specified, the returned string will be
'    '             filled up with spaces.
'    '
'    If Len(padChar) Then
'        PadLeft = StrVal
'        If Len(StrVal) <= totalWidth Then _
'            PadLeft = String$(totalWidth - Len(StrVal), padChar) & PadLeft
'    Else
'        PadLeft = Space$(totalWidth)
'        RSet PadLeft = StrVal
'    End If
'
'End Function
'
'Public Function PadRight(StrVal As String, _
'                         ByVal totalWidth As Long, _
'                         Optional ByVal padChar As String) As String
'
'    ' Returns the String StrVal with the specified length.
'    ' totalWidth: the length of the returned string
'    '             if totalWidth is smaller then the length of StrVal then
'    '             StrVal will be returned
'    ' padChar:    on the right hand side it will be filed up with padChar
'    '             if padChar is not specified, the returned string will be
'    '             filled up with spaces.
'    '
'    If Len(padChar) Then
'        PadRight = StrVal
'        If Len(StrVal) <= totalWidth Then _
'            PadRight = PadRight & String$(totalWidth - Len(StrVal), padChar)
'    Else
'        PadRight = Space$(totalWidth)
'        LSet PadRight = StrVal
'    End If
'
'End Function

Public Function Remove(s As String, ByVal startIndex As Long, Optional ByVal Count As Long = -1) As String
    'Remove(Int32, Int32)
    'Gibt eine neue Zeichenfolge zurück, in der eine bestimmte Anzahl von Zeichen in
    'der aktuellen Instanz, beginnend an einer angegebenen Position, gelöscht wurden.
    'Remove (Int32)
    'Gibt eine neue Zeichenfolge zurück, in der alle Zeichen in der aktuellen Instanz,
    'beginnend an einer angegebenen Position und sich über die letzte Position
    'fortsetzend, gelöscht wurden.
    
End Function

'Public Function Replace() As String
'    '
'End Function

Public Function StartsWith(s As String, ByVal Value As String) As Boolean
    '
End Function

Public Function Substring(s As String, Optional ByVal startIndex As Long, Optional ByVal length As Long) As String
    '
End Function

Public Function ToCharArray(ByVal startIndex As Long, ByVal length As Long) As Integer()
    '
End Function
