Attribute VB_Name = "MString"
Option Explicit 'Zeilen: 129; 2022.01.06 Zeilen: 336; 2022.11.01 Zeilen: 625;
'https://learn.microsoft.com/de-de/cpp/text/how-to-convert-between-various-string-types?view=msvc-170

'https://de.wikipedia.org/wiki/Byte_Order_Mark
Public Enum EByteOrderMark
    bom_UTF_8 = &HBFBBEF        '                     '     239 187 191     '  ï»¿           ' [4]
    bom_UTF_16_BE = &HFFFE&     ' Big Endian Motorola '         254 255     '   þÿ
    bom_UTF_16_LE = &HFEFF&     ' little endian Intel '         255 254     '   ÿþ
    bom_UTF_32_BE = &HFFFE0000  ' Big Endian Motorola '   0   0 254 255     ' ??þÿ
    bom_UTF_32_LE = &HFEFF      ' little endian Intel ' 255 254   0   0     ' ÿþ??
    bom_UTF_7 = &H762F2B        '                     '      43  47 118
                                ' und ein Zeichen aus: [ 56 | 57 | 43 | 47 ]
                                ' und ein Zeichen aus: [ 38 | 39 | 2B | 2F ]           ' [5]
                                ' + / v und ein Zeichen aus:  [  8 |  9 |  + |  / ]
    bom_UTF_1 = &H4C64F7        '                     '     247 100  76     ' ÷dL
    bom_UTF_EBCDIC = &H736673DD '                     ' 221 115 102 115     ' Ýsfs
    bom_SCSU = &HFFFE0E         '                     '      14 254 255     ' ?þÿ            ' [6]
                                ' (von anderen möglichen Bytefolgen wird abgeraten)
    bom_BOCU_1 = &H28EEFB       '                     '     251 238  40
                                ' optional gefolgt von FF                              ' [7]
                                ' optional gefolgt von 255     ûî
                                ' optional gefolgt von          ÿ
    bom_GB_18030 = &H33953184   '               ' 132  49 149  51     ' „1•3
End Enum

'Maybe we need an enum Encoding
Private Const CP_UTF8 As Long = 65001

'Public Enum ECodePage
'
'End Enum

'#Const Unicode = 1

'#If VBA7 = 0 Then
'    Private Enum LongPtr
'        [_]
'    End Enum
'#End If
#If VBA7 Then
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal pDst As LongPtr, ByVal pSrc As LongPtr) As Long
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal wType As Long) As Long
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)
#Else
    'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
    Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
    'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar
    Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare Function lstrcpyW Lib "kernel32" (ByVal pDst As LongPtr, ByVal pSrc As LongPtr) As Long
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal wType As Long) As Long
    Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)
#End If
Private Type TCur
    Value As Currency
End Type
Private Type TLong2
    Lo As Long
    Hi As Long
End Type

#If False Then
    Value
#End If
'VB does automatic in and out Ansi/Unicode conversion when calling winapi-functions with parameters of type String
'you can simulate this behaviour by using StrPtrWA in the call and WACorr afterwards
'Public Function StrPtrWA(ByRef s_inout As String) As LongPtr
'#If Unicode Then
'    StrPtrWA = StrPtr(s_inout)
'#Else
'    s_inout = StrConv(s_inout, vbFromUnicode)
'    StrPtrWA = StrPtr(s_inout)
'#End If
'End Function

'Public Sub WACorr(ByRef s_inout As String)
'#If Unicode Then
'    '
'#Else
'    s_inout = StrConv(s_inout, vbUnicode)
'#End If
'End Sub

Public Function Trim0(ByVal s As String) As String
    Trim0 = VBA.Strings.Trim$(Left$(s, lstrlenW(ByVal StrPtr(s))))
End Function

'Private Function Ptr2String(ByVal pStr As LongPtr) As String
'    If pStr = 0 Then Exit Function
'    Dim l As Long: l = lstrlenW(pStr) * 2& - 1&
'    If l > 0& Then
'        ReDim bBuffer(l) As Byte
'        RtlMoveMemory bBuffer(0), ByVal pStr, l
'        CoTaskMemFree lpStrPointer
'        Ptr2String = bBuffer
'    End If
'End Function

Public Function PtrToString(ByVal pStr As LongPtr, Optional ByVal sLen As Long) As String
    If (pStr = 0) Then Exit Function
    If sLen <= 0 Then sLen = lstrlenW(pStr)
    PtrToString = Space$(sLen)
    lstrcpyW StrPtr(PtrToString), pStr
    'CoTaskMemFree pStr ' is das so immer richtig?
'#If defUnicode Then
'    'ist es dann schon der richtige String?
'    'MsgBox PtrToString
'#Else
'    PtrToString = Left$(StrConv(PtrToString, vbUnicode), num1)
'#End If
End Function
Public Function PtrToStringCo(ByVal pStr As LongPtr, Optional ByVal sLen As Long) As String
    If (pStr = 0) Then Exit Function
    If sLen <= 0 Then sLen = lstrlenW(pStr)
    PtrToStringCo = Space$(sLen)
    lstrcpyW StrPtr(PtrToStringCo), pStr
    CoTaskMemFree pStr ' is das so immer richtig?
'#If defUnicode Then
'    'ist es dann schon der richtige String?
'    'MsgBox PtrToString
'#Else
'    PtrToString = Left$(StrConv(PtrToString, vbUnicode), num1)
'#End If
End Function

Public Function IsHex(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        Select Case Asc(Mid(s, i, 1))
        Case 48 To 57:  ' 0 - 9 OK weiter
        Case 65 To 70:  ' A - F OK weiter
        Case 97 To 102: ' a - f OK weiter
        Case Else: Exit Function
        End Select
    Next
    IsHex = True
End Function

'Dim fnam As String: fnam = Left(lpElfe.lfFontName, lstrlenW(lpElfe.lfFontName(0)))

'Private Function CreateGUID() As String
'    Dim G As Guid, GuidByt As Long, l As Long, GuidStr As String, Buffer() As Byte
'    If UuidCreate(G) <> RPC_S_UUID_NO_ADDRESS Then
'        If UuidToString(G, GuidByt) = RPC_S_OK Then
'            l = lstrlen(GuidByt)
'            ReDim Buffer(l - 1) As Byte
'            RtlMoveMemory Buffer(0), ByVal GuidByt, l
'            RpcStringFree GuidByt
'            GuidStr = StrConv(Buffer, vbUnicode)
'            CreateGUID = UCase$(GuidStr)
'        End If
'    End If
'End Function

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

Public Function RecursiveReplaceSL(ByVal Expression As String, ByVal Find As String, ByVal Replace As String, Optional ByVal Start As Long = 1, Optional ByVal Length As Long = -1) As String
    'Uses RecursiveReplace to replace "Find" by "Replace" in a part of "Expression" that starts with "Start" with the length of "Length"
    'check input parameters return early if necessary
    If Length < 0 And Start = 1 Then RecursiveReplaceSL = RecursiveReplace(Expression, Find, Replace): Exit Function
    Dim le As Long: le = Len(Expression)
    If Start < 1 Or le < Start Then Exit Function 'return nothing
    If Length < 1 Or le < Start + Length Then Length = le - Start + 1
    
    Dim sl As String: sl = Left$(Expression, Start - 1)
    Dim sm As String: sm = Mid$(Expression, Start, Length)
    Dim sr As String: sr = Mid$(Expression, Start + Length)
    sm = RecursiveReplace(sm, Find, Replace)
    RecursiveReplaceSL = sl & sm & sr
    'same but shorter and less noise:
    'RecursiveReplaceSL = Left$(Expression, Start - 1) & RecursiveReplace(Mid$(Expression, Start, Length)) & Mid$(Expression, Start, Length)
End Function

'used by StringClean in MIBANUtil
Public Function ReplaceAll(ByVal Expression As String, Find As String, Replace As String) As String
    Dim i As Integer
    For i = 1 To Len(Expression)
        Expression = VBA.Replace(Expression, Mid(Find, i, 1), Replace)
    Next
    ReplaceAll = Expression
End Function


'Converters to or from String
'Bool
Public Function BoolToYesNo(ByVal b As Boolean) As String
    BoolToYesNo = IIf(b, " Ja ", "Nein")
End Function

Public Function StrToBol(ByVal s As String) As Boolean
    s = UCase$(Trim$(s))
    If s = "yes" Then StrToBol = True: Exit Function
    If s = "ja" Then StrToBol = True: Exit Function
    If s = "ok" Then StrToBol = True: Exit Function
    If s = "1" Then StrToBol = True: Exit Function
    If s = "-1" Then StrToBol = True: Exit Function
    If s = "wahr" Then StrToBol = True: Exit Function
    If s = "true" Then StrToBol = True: Exit Function
    StrToBol = CBool(s)
End Function

'Private Function StrToBol(StrVal As String) As Boolean
'    If (StrComp(StrVal, "0", vbTextCompare) = 0) Or _
'       (StrComp(StrVal, "false", vbTextCompare) = 0) Or _
'       (StrComp(StrVal, "falsch", vbTextCompare) = 0) Or _
'       (StrComp(StrVal, "nein", vbTextCompare) = 0) Then
'        StrToBol = False
'    'ElseIf (StrComp(StrVal, vbNullString) = 0) Or _
'           (StrComp(StrVal, "1") = 0) Or _
'           (StrComp(StrVal, "-1") = 0) Or _
'           (StrComp(StrVal, "true") = 0) Or _
'           (StrComp(StrVal, "wahr") = 0) Or _
'           (StrComp(StrVal, "ja") = 0) Then
'    Else
'        StrToBol = True
'    End If
'End Function

Public Function BolToStr(ByVal b As Boolean) As String
    If b Then BolToStr = "True" Else BolToStr = "False"
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

Public Function Hex16(ByVal Value As Currency) As String
    Dim tc As TCur:  tc.Value = Value
    Dim tl As TLong2: LSet tl = tc
    Hex16 = Hex8(tl.Hi) & Hex8(tl.Lo)
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
Public Function IndexOf(s As String, ByVal Value As String, Optional ByVal startIndex As Long = 0, Optional ByVal Count As Long = -1, Optional ByVal compare As VbCompareMethod = vbBinaryCompare) As Long
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
    IndexOf = InStr(1, v, Value, compare) - 1
    If IndexOf > 0 Then IndexOf = startIndex + IndexOf - 1
End Function

Public Function Insert(s As String, ByVal startIndex As Long, ByVal Value As String) As String
    Insert = Left(s, startIndex) & Value & Mid(s, startIndex)
End Function

Public Function LastIndexOf(s As String, Value As String, ByVal startIndex As Long, ByVal Count As Long, Optional ByVal compare As VbCompareMethod = vbBinaryCompare) As Long
    Dim pos As Long: pos = InStrRev(s, Value, startIndex, compare)
    LastIndexOf = pos
End Function

Public Function GetDecimalSeparator() As String
    Dim d As Double, s As String
    s = Format(d, "0.0")
    GetDecimalSeparator = Mid(s, 2, 1)
End Function

Function PadLeft(this As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String) As String
    
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
        If Len(this) < totalWidth Then
            PadLeft = Space$(totalWidth)
            RSet PadLeft = this
        Else
            PadLeft = this
        End If
    End If
End Function

Function PadRight(this As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String) As String
    
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
        If Len(this) < totalWidth Then
            PadRight = Space$(totalWidth)
            LSet PadRight = this
        Else
            PadRight = this
        End If
    End If
End Function

'Public Function PadLeft2(StrVal As String, _
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
'        PadLeft2 = StrVal
'        If Len(StrVal) <= totalWidth Then _
'            PadLeft2 = String$(totalWidth - Len(StrVal), padChar) & PadLeft2
'    Else
'        PadLeft2 = Space$(totalWidth)
'        RSet PadLeft2 = StrVal
'    End If
'
'End Function
''
'Public Function PadRight2(StrVal As String, _
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
'        PadRight2 = StrVal
'        If Len(StrVal) <= totalWidth Then _
'            PadRight2 = PadRight2 & String$(totalWidth - Len(StrVal), padChar)
'    Else
'        PadRight2 = Space$(totalWidth)
'        LSet PadRight2 = StrVal
'    End If
'
'End Function
'
'Dim s As String = "Dies ist ein String"
'Remove(startIndex, count)
's = s.Remove(-1)     ' "" und Fehlermeldung
's = s.Remove(0)      ' ""
's = s.Remove(10)     ' "Dies ist e"
's = s.Remove(19)     ' "Dies ist ein String"
's = s.Remove(20)     ' "Dies ist ein String"  und Fehlermeldung

's = s.Remove(-1, 0)  ' "" und Fehlermeldung
's = s.Remove(0, 0)   ' "Dies ist ein String"
's = s.Remove(10, 0)  ' "Dies ist ein String"
's = s.Remove(19, 0)  ' "Dies ist ein String"
's = s.Remove(20, 0)  ' "Dies ist ein String"  und Fehlermeldung

's = s.Remove(-1, 10) ' "" und Fehlermeldung
's = s.Remove(0, 10)  ' "in String"
's = s.Remove(1, 10)  ' "Dn String"
's = s.Remove(7, 6)   ' "Dies isString"
's = s.Remove(10, 9)  ' "Dies ist e"
's = s.Remove(10, 10) ' "Dies ist e" und Fehlermeldung

's = s.Remove(-1, 19) ' "" und Fehlermeldung
's = s.Remove(0, 19)  ' ""
's = s.Remove(1, 19)  ' "" und Fehlermeldung
Public Function Remove(s As String, ByVal startIndex As Long, Optional ByVal Count As Long = -1) As String
    'Remove(Int32, Int32)
    'Gibt eine neue Zeichenfolge zurück, in der eine bestimmte Anzahl von Zeichen in
    'der aktuellen Instanz, beginnend an einer angegebenen Position, gelöscht wurden.
    'Remove (Int32)
    'Gibt eine neue Zeichenfolge zurück, in der alle Zeichen in der aktuellen Instanz,
    'beginnend an einer angegebenen Position und sich über die letzte Position
    'fortsetzend, gelöscht wurden.
    'ist startindex 1-basiert?
    'If startIndex = 0 And Count = -1 Then
    'Dim pos As Long: pos = Len(s) - startIndex
    Dim L As Long: L = Len(s)
    If Count < 0 Then
        If startIndex < 0 Then
            Remove = ""
            'Error message
            Exit Function
        End If
        If startIndex = 0 Then
            Remove = ""
            Exit Function
        End If
        If startIndex < L Then
            Remove = Left$(s, startIndex)
            Exit Function
        End If
        If startIndex = L Then
            Remove = s
            Exit Function
        End If
        Remove = s
        'Error message
        Exit Function
    End If
    If Count = 0 Then
        If startIndex < 0 Then
            Remove = ""
            'Error message
            Exit Function
        End If
        If startIndex = 0 Then
            Remove = s
            Exit Function
        End If
        If startIndex < L Then
            Remove = s
            Exit Function
        End If
        If startIndex = L Then
            Remove = s
            Exit Function
        End If
        Remove = s
        'Error message
        Exit Function
    End If
    If Count < L Then
        If startIndex < 0 Then
            Remove = ""
            'Error message
            Exit Function
        End If
        If startIndex = 0 Then
            Remove = Mid$(s, Count + 1)
            Exit Function
        End If
        If startIndex < L Then
            If startIndex + Count < L Then
                Remove = Left(s, startIndex) & Mid(s, startIndex + Count + 1)
                Exit Function
            End If
            If startIndex + Count = L Then
                Remove = Left(s, startIndex)
                Exit Function
            End If
            If L < startIndex + Count Then
                Remove = Left(s, startIndex)
                'Error message
                Exit Function
            End If
        End If
        If startIndex = L Then
            Remove = s
            'Error message
            Exit Function
        End If
        Remove = s
        'Error message
        Exit Function
    End If
    If Count = L Then
        If startIndex < 0 Then
            Remove = ""
            'Error message
            Exit Function
        End If
        If startIndex = 0 Then
            Remove = "" 'Mid$(s, Count + 1)
            Exit Function
        End If
        If startIndex < L Then
            Remove = Left$(s, startIndex)
            'Error message
            Exit Function
        End If
    End If
    If L < Count Then
        If startIndex < 0 Then
            Remove = ""
            'Error message
            Exit Function
        End If
        If startIndex = 0 Then
            Remove = "" 'Mid$(s, Count + 1)
            'Error message
            Exit Function
        End If
        If startIndex < L Then
            Remove = Left(s, startIndex)
            'Error message
            Exit Function
        End If
    End If
End Function

Public Function StartsWith(s As String, ByVal Value As String) As Boolean
    StartsWith = Left$(s, Len(Value)) = Value
End Function

Public Function Substring(s As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long) As String
    Substring = Mid(s, startIndex, Length)
End Function

Public Function ToCharArray(s As String, ByVal startIndex As Long, ByVal Length As Long) As Integer()
    ReDim CharArray(0 To Length - 1) As Integer
    lstrcpyW VarPtr(CharArray(0)), StrPtr(Mid$(s, startIndex, Length))
    ToCharArray = CharArray
End Function

Public Function SArray(ParamArray strArr()) As String()
    ReDim sa(0 To UBound(strArr)) As String
    Dim i As Long: For i = 0 To UBound(strArr): sa(i) = strArr(i): Next
    SArray = sa
End Function

Public Function EByteOrderMark_Parse(ByVal Value As Long) As EByteOrderMark
    
    Dim e  As Long
    
    e = Value
    If e = EByteOrderMark.bom_UTF_32_BE Or _
       e = EByteOrderMark.bom_GB_18030 Or _
       e = EByteOrderMark.bom_UTF_EBCDIC Or _
       e = EByteOrderMark.bom_UTF_32_LE Then EByteOrderMark_Parse = e: Exit Function
    
    e = Value And &HFFFFFF
    If e = EByteOrderMark.bom_SCSU Or _
       e = EByteOrderMark.bom_UTF_8 Or _
       e = EByteOrderMark.bom_BOCU_1 Or _
       e = EByteOrderMark.bom_UTF_1 Then EByteOrderMark_Parse = e: Exit Function
       
    If e = EByteOrderMark.bom_UTF_7 Then
        e = Value \ 2 ^ 24 'shiftright 24 bits
        If e = &H38 Or e = &H39 Or e = &H2B Or e = &H2F Then _
                    EByteOrderMark_Parse = EByteOrderMark.bom_UTF_7: Exit Function
    End If
    
    e = Value And &HFFFF&
    If e = EByteOrderMark.bom_UTF_16_BE Or _
       e = EByteOrderMark.bom_UTF_16_LE Then EByteOrderMark_Parse = e: Exit Function
    
End Function

Public Function EByteOrderMark_ToStr(ByVal Value As EByteOrderMark) As String
    Dim s As String
    Dim e As EByteOrderMark
    Select Case Value
    Case e = bom_BOCU_1:     s = "bom_BOCU_1"
    Case e = bom_SCSU:       s = "bom_SCSU"
    Case e = bom_UTF_1:      s = "bom_UTF_1"
    Case e = bom_UTF_16_BE:  s = "bom_UTF_16_BE"
    Case e = bom_UTF_16_LE:  s = "bom_UTF_16_LE"
    Case e = bom_UTF_32_BE:  s = "bom_UTF_32_BE"
    Case e = bom_UTF_7:      s = "bom_UTF_7"
    Case e = bom_UTF_8:      s = "bom_UTF_8"
    Case e = bom_UTF_EBCDIC: s = "bom_UTF_EBCDIC"
    End Select
    EByteOrderMark_ToStr = s
End Function

Public Function ConvertFromUTF8(ByRef Source() As Byte) As String
    'All credits for this function are going to Philipp Stephani from ActiveVB
    'http://www.activevb.de/rubriken/faq/faq0155.html
    Dim Size    As Long:       Size = UBound(Source) - LBound(Source) + 1
    Dim pSource As LongPtr: pSource = VarPtr(Source(LBound(Source)))
    Dim Length  As Long:     Length = MultiByteToWideChar(CP_UTF8, 0, pSource, Size, 0, 0)
    Dim Buffer  As String:   Buffer = Space$(Length)
    MultiByteToWideChar CP_UTF8, 0, pSource, Size, StrPtr(Buffer), Length
    ConvertFromUTF8 = Buffer
    
End Function

Public Property Get App_EXEName() As String
#If VBA6 Or VBA7 Then
    App_EXEName = Application.Name
#Else
    App_EXEName = App.EXEName
#End If
End Property

Public Function GetGreekAlphabet() As String
    Dim s As String
    Dim i As Long
    Dim alp As Long: alp = 913 'der Große griechische Buchstabe Alpha
    For i = alp To alp + 24
        s = s & ChrW(i)
    Next
    s = s & " "
    alp = alp + 32             'der Kleine griechische Buchstabe alpha
    For i = alp To alp + 24
        s = s & ChrW(i)
    Next
    GetGreekAlphabet = s
End Function

Public Function MsgBoxW(Prompt, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title) As VbMsgBoxResult
'Public Function MsgBoxW(Prompt, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title, Optional Helpfile, Optional Context) As VbMsgBoxResult
    Title = IIf(IsMissing(Title), App_EXEName, CStr(Title))
    MsgBoxW = MessageBoxW(0, StrPtr(Prompt), StrPtr(Title), Buttons)
End Function


