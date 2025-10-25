Attribute VB_Name = "MString"
Option Explicit 'lines 129; 2022.01.06 lines 336; 2022.11.01 lines 625; 2024-06-24 lines 1595; 2024-07-17 lines 1948; 2024-08-25 lines: 2102; 2025-09-13 lines: 2378;
'For using this module you also must include:
'..\Ptr_Pointers\MPtr.bas
'..\Math\MMath.bas

'https://learn.microsoft.com/de-de/cpp/text/how-to-convert-between-various-string-types?view=msvc-170
'https://de.wikipedia.org/wiki/Byte_Order_Mark
Public Enum EByteOrderMark

    bom_None = 0
                                '    Zeichenfolge:        4  3  2  1
    bom_UTF_8 = &HBFBBEF        '    bom_UTF_8      = &H    BF BB EF  '                     '     239 187 191     ' ï»¿     ' [4]
    bom_UTF_16_BE = &HFFFE      '    bom_UTF_16_BE  = &H       FF FE  ' Big Endian Motorola '         254 255     ' þÿ
    bom_UTF_16_LE = &HFEFF      '    bom_UTF_16_LE  = &H       FE FF  ' little endian Intel '         255 254     ' ÿþ
    bom_UTF_32_BE = &HFFFE0000  '    bom_UTF_32_BE  = &H FF FE 00 00  ' Big Endian Motorola '   0   0 254 255     ' ??þÿ
    bom_UTF_32_LE = &HFEFF&     '    bom_UTF_32_LE  = &H 00 00 FE FF  ' little endian Intel ' 255 254   0   0     ' ??ÿþ
    bom_UTF_7 = &H762F2B        '    bom_UTF_7      = &H    76 2F 2B  '                     '      43  47 118     ' +/v     ' und ein Zeichen aus:  [  8 |  9 |  + |  / ]
                                '      und aus: [&H38|&H39|&H2B|&H2F] ' und ein Zeichen aus: [ 56| 57| 43| 47]              ' [5]
    bom_UTF_1 = &H4C64F7        '    bom_UTF_1      = &H    4C 64 F7  '                     '     247 100  76     ' ÷dL
    bom_UTF_EBCDIC = &H736673DD '    bom_UTF_EBCDIC = &H 73 66 73 DD  '                     ' 221 115 102 115     ' Ýsfs
    bom_SCSU = &HFFFE0E         '    bom_SCSU       = &H    FF FE 0E  '                     '      14 254 255     ' ?þÿ     ' [6]   (von anderen möglichen Bytefolgen wird abgeraten)
    bom_BOCU_1 = &H28EEFB       '    bom_BOCU_1     = &H    28 EE FB  '                     '     251 238  40     ' ûî(     ' [7]   optional gefolgt von ÿ
                                '    optional gefolgt von FF          ' optional gefolgt von              255
    bom_GB_18030 = &H33953184   '    bom_GB_18030   = &H 33 95 31 84  '                     ' 132  49 149  51     ' „1•3

End Enum

#If VBA7 Then
Public Enum ShiftConstants
     vbShiftMask = 1
     vbCtrlMask = 2
     vbAltMask = 4
End Enum
#End If

'Maybe we need an enum Encoding
Private Const CP_ACP        As Long = 0     ' The system default Windows ANSI code page.
Private Const CP_OEMCP      As Long = 1     ' The current system OEM code page.
Private Const CP_MACCP      As Long = 2     ' The current system Macintosh code page.
Private Const CP_THREAD_ACP As Long = 3     ' The Windows ANSI code page for the current thread.
Private Const CP_SYMBOL     As Long = 42    ' Symbol code page (42).
Private Const CP_WINUNICODE As Long = 1200
Private Const CP_UTF7       As Long = 65000 ' UTF-7. Use this value only when forced by a 7-bit transport mechanism. Use of UTF-8 is preferred.
Private Const CP_UTF8       As Long = 65001 ' UTF-8.

Public Enum ETextEncoding
    Text_ASCIIEncoding = 0
    Text_UnicodeEncoding = 1200     ' UCS2 or UTF16 (2 bytes per char)
    Text_UTF32Encoding              '               (4 bytes per char)
    Text_UTF7Encoding = 65000
    Text_UTF8Encoding = 65001
End Enum

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
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
    'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar
    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
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

'for Base64-encoding:
                        '65          -        90, 97         -         122, 48-57, 43, 47
Private Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private m_B64_IsInitialized As Boolean
Private B64() As Byte
Private RevB64() As Byte

Private Const m_ValidateMsg As String = "Please give a valid <datatype> value: <value>" & vbCrLf & "OK=improve your input; Cancel=last value"

'vbvartype here is used for vbtypeids also, this are extensions to vbtypeids maybe we should have an enum?
Private Const vbHex   As Long = &H10000
Private Const vbOct   As Long = &H20000
Private Const vbBin   As Long = &H40000
Private Const vbIdent As Long = &H80000

Public DecimalSeparator As String
Public CurrencySymbol   As String
Private m_isInitialized As Boolean
Private m_VBOperators() As String

Public Function Init()
    If Not m_isInitialized Then
        m_VBOperators() = Split(" Or , Xor , And , + , - , * , / ", ",")
        CurrencySymbol = "€"
        m_isInitialized = True
    End If
End Function

Public Function FormatMByte(CountElements As Long, vt As VbVarType) As String
    Dim bpe As Long
    Select Case vt
    Case vbByte:                        bpe = 1
    Case vbInteger, vbString:           bpe = 2 ' 2 bytes per character
    Case vbLong, vbSingle:              bpe = 4
    Case vbCurrency, vbDouble, vbDate:  bpe = 8
    Case vbDecimal, vbVariant:          bpe = 16
    End Select
    FormatMByte = Format(CountElements * bpe / 1024 / 1024, "0.000") & " MByte"
End Function

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
'
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
    If sLen <= 0 Then sLen = lstrlenW(pStr) '- 1
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

Public Function IsOct(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        Select Case Asc(Mid(s, i, 1))
        Case 48 To 55:  ' 0 - 7 OK weiter
        Case Else: Exit Function
        End Select
    Next
    IsOct = True
End Function

Public Function IsBin(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        Select Case AscW(Mid(s, i, 1))
        Case 48, 49: ' 0 oder 1 OK weiter
        Case Else: Exit Function
        End Select
    Next
    IsBin = True
End Function

Public Function IsHexPrefix(ByVal s As String) As Boolean
    IsHexPrefix = Left(Trim(s), 2) = "&H"
End Function

Public Function IsOctPrefix(ByVal s As String) As Boolean
    IsOctPrefix = Left(Trim(s), 2) = "&O"
End Function

Public Function IsBinPrefix(ByVal s As String) As Boolean
    IsBinPrefix = Left(Trim(s), 2) = "&B"
End Function

Public Function IsExpression(ByVal s As String) As Boolean
    If Not m_isInitialized Then MString.Init
    IsExpression = MString.ContainsOneOf(s, m_VBOperators)
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
    'Recursive replace/delete multi WhiteSpaces WS
    DeleteMultiWS = Trim$(s)
    If InStr(1, s, "  ") = 0 Then Exit Function
    DeleteMultiWS = Replace(s, "  ", " ")
    DeleteMultiWS = DeleteMultiWS(DeleteMultiWS)
End Function

Public Function DeleteCRLF(s As String, Optional replacewith As String = " ") As String
    'Recursive replace/delete multi crlfs
    DeleteCRLF = Trim$(s)
    If InStr(1, s, vbLf) = 0 Then Exit Function
    If InStr(1, s, vbCr) = 0 Then Exit Function
    DeleteCRLF = Replace(Replace(Replace(s, vbCrLf, replacewith), vbLf, replacewith), vbCr, replacewith)
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

' v ############################## v '    TryParse & ToStr Functions    ' v ############################## v '
'Converters to or from String
'Bool
Public Function BoolToYesNo(ByVal b As Boolean) As String
    'BoolToYesNo = IIf(b, "yes ", " no ")
    BoolToYesNo = IIf(b, " Ja ", "Nein")
End Function

Public Function CBol(ByVal s As String) As Boolean
Try: On Error GoTo Catch
    s = UCase$(Trim$(s))
    If s = "YES" Then CBol = True: Exit Function
    If s = "JA" Then CBol = True: Exit Function
    If s = "OK" Then CBol = True: Exit Function
    If s = "1" Then CBol = True: Exit Function
    If s = "-1" Then CBol = True: Exit Function
    If s = "WAHR" Then CBol = True: Exit Function
    If s = "TRUE" Then CBol = True: Exit Function
    CBol = CBool(s)
Catch:
End Function

Public Function StrToBol(ByVal s As String) As Boolean
    StrToBol = CBol(s)
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

Public Function Byte_TryParse(ByVal Value As String, ByRef Value_out As Byte) As Boolean
Try: On Error GoTo Catch
    Value_out = CByte(Value)
    Byte_TryParse = True
    Exit Function
Catch:
End Function

Public Function Byte_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As Byte) As Boolean
    Byte_TryParseMess = Byte_TryParse(Value, Value_inout)
    If Not Byte_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbByte))
End Function

Public Function Byte_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As Byte) As String
    bIsOK_out = Byte_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        Byte_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbByte))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        Byte_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Double value
    Byte_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function

Public Function Integer_TryParse(ByVal Value As String, ByRef Value_out As Integer) As Boolean
Try: On Error GoTo Catch
    Value = Trim(Value)
    If Right(Value, 1) = "%" Then Value = Left(Value, Len(Value) - 1)
    Value_out = CInt(Value)
    Integer_TryParse = True
    Exit Function
Catch:
End Function

Public Function Integer_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As Integer) As Boolean
    Integer_TryParseMess = Integer_TryParse(Value, Value_inout)
    If Not Integer_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbInteger))
End Function

Public Function Integer_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As Integer) As String
    bIsOK_out = Integer_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        Integer_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbInteger))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        Integer_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Double value
    Integer_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function

Public Function Boolean_TryParse(ByVal Value As String, ByRef Value_out As Boolean) As Boolean
Try: On Error GoTo Catch
    Value = UCase$(Trim$(Value))
    Select Case Value
    Case "YES", "JA", "OK", "1", "-1", "WAHR", "TRUE":           Value = True
    Case "NO", "NEIN", "NOTOK", "0", "FALSCH", "FALSE", "WRONG": Value = False
    End Select
    Value_out = CBool(Value)
    Boolean_TryParse = True
    Exit Function
Catch:
End Function

Public Function Boolean_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As Boolean) As Boolean
    Boolean_TryParseMess = Boolean_TryParse(Value, Value_inout)
    If Not Boolean_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbBoolean))
End Function

Public Function Boolean_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As Boolean) As String
    bIsOK_out = Boolean_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        Boolean_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbBoolean))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        Boolean_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Double value
    Boolean_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function

Public Function Long_TryParse(ByVal Value As String, ByRef Value_out As Long) As Boolean
Try: On Error GoTo Catch
    Value = Trim(Value)
    If Right(Value, 1) = "&" Then Value = Left(Value, Len(Value) - 1)
    Value_out = CLng(Value)
    Long_TryParse = True
    Exit Function
Catch:
End Function

Public Function Long_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As Long) As Boolean
    Long_TryParseMess = Long_TryParse(Value, Value_inout)
    If Not Long_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbLong))
End Function

Public Function Long_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As Long) As String
    bIsOK_out = Long_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        Long_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbLong))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        Long_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Double value
    Long_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function

#If VBA7 Then
Public Function LongLong_TryParse(ByVal Value As String, ByRef Value_out As LongLong) As Boolean
Try: On Error GoTo Catch
    Value_out = CLng(Value)
    Long_TryParse = True
    Exit Function
Catch:
End Function

Public Function LongLong_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As LongLong) As Boolean
    LongLong_TryParseMess = LongLong_TryParse(Value, Value_inout)
    If Not LongLong_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbLongLong))
End Function

Public Function LongLong_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As LongLong) As String
    bIsOK_out = LongLong_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        LongLong_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbLongLong))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        LongLong_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Double value
    LongLong_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function
#End If

Public Function Single_TryParse(ByVal Value As String, ByRef Value_out As Single) As Boolean
Try: On Error GoTo Catch
    Value = Trim(Value)
    If Right(Value, 1) = "!" Then Value = Left(Value, Len(Value) - 1)
    Value = Replace(Value, ",", ".")
    If Not IsNumeric(Value) Then Exit Function
    Value_out = CSng(Val(Value)) 'hey was das für ne scheiße, warum nicht Val() ????
    'Value_out = CSng(Value)     'hey was das für ne scheiße, warum nicht Val() ????
    Single_TryParse = True
    Exit Function
Catch:
End Function

Public Function Single_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As Single) As Boolean
    Single_TryParseMess = Single_TryParse(Value, Value_inout)
    If Not Single_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbSingle))
End Function

Public Function Single_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As Single) As String
    bIsOK_out = Single_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        Single_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbSingle))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        Single_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Double value
    Single_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function

Public Function Double_TryParse(ByVal Value As String, ByRef Value_inout As Double) As Boolean
Try: On Error GoTo Catch
    Value = Trim(Value)
    If Len(Value) = 0 Then Exit Function
    Value = Replace(Value, "­", "-") 'replace &HC2AD with &H2D
    Value = Replace(Value, ",", ".") 'for using the function Val()
    If IsNumeric(Value) Then
        'If Right(Value, 1) = "#" Then Value = Left(Value, Len(Value) - 1)
        Value_inout = Val(Value)
        Double_TryParse = True: Exit Function
    Else
        If StrComp(Value, "1.#QNAN") = 0 Then
            MMath.GetNaN Value_inout:           Double_TryParse = True: Exit Function
        ElseIf StrComp(Value, "1.#INF") = 0 Then
            Value_inout = MMath.GetINF:         Double_TryParse = True: Exit Function
        ElseIf StrComp(Value, "-1.#INF") = 0 Then
            Value_inout = MMath.GetINF(-1):     Double_TryParse = True: Exit Function
        ElseIf StrComp(Value, "-1.#IND") = 0 Then
            MMath.GetINDef Value_inout:         Double_TryParse = True: Exit Function
        End If
    End If
Catch:
End Function

Public Function Double_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As Double) As Boolean
    Double_TryParseMess = Double_TryParse(Value, Value_inout)
    If Not Double_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbDouble))
End Function

Public Function Double_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As Double) As String
    bIsOK_out = Double_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        Double_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbDouble))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        Double_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Double value
    Double_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function

'Public Function Double_TryParseValue(ByVal OldValue As String, ByVal NewValue As String, ByVal mess As String, ByRef Value_inout As Double) As String
    'Double_TryParseMess = Double_TryParse(Value, Value_inout)
    'hmm until now we just let the user say "OK" actually there is a better approach:
    'lets say the message is "Please give a valid value: "
    'then the use can agree and you give em the wrong value back so the user gets tha chance to remedy his fault
    'so the user says "OK" I will do it better next time so please give me back my editings
    'or the user says "Oh no this was a complete mess please give me the old value by clicking "Cancel"
    'so you givbe the old value
    'If Not Double_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbDouble))

Public Function Date_TryParse(ByVal Value As String, ByRef Value_out As Date) As Boolean
Try: On Error GoTo Catch
    Value_out = CDate(Value)
    Date_TryParse = True
Catch:
End Function

Public Function Date_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As Date) As Boolean
    Date_TryParseMess = Date_TryParse(Value, Value_inout)
    If Not Date_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbDate))
End Function

Public Function Date_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As Date) As String
    bIsOK_out = Date_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        Date_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbDate))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        Date_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Date value
    Date_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function

Public Function Currency_TryParse(ByVal Value As String, ByRef Value_out As Currency) As Boolean
Try: On Error GoTo Catch
    Value = Trim(Value)
    If Right(Value, 1) = "@" Then Value = Left(Value, Len(Value) - 1)
    Value = Replace(Value, "­", "-") 'replace &HC2AD with &H2D
    Dim ds As String: ds = GetDecimalSeparator
    Value = Replace(Value, ",", ds)
    Value = Replace(Value, ".", ds)
    Value_out = CCur(Value)
    Currency_TryParse = True
Catch:
End Function

Public Function Currency_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout As Currency) As Boolean
    Currency_TryParseMess = Currency_TryParse(Value, Value_inout)
    If Not Currency_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbCurrency))
End Function

Public Function Currency_TryParseValidate(ByVal NewValue As String, ByVal mess As String, ByVal sFormat As String, ByRef bIsOK_out As Boolean, ByRef OldValueIn_NewValueOut As Currency) As String
    bIsOK_out = Currency_TryParse(NewValue, OldValueIn_NewValueOut)
    If bIsOK_out Then
        'OldValueIn_NewValueOut now has changed to the new value because user gave a valid value in NewValue
        Currency_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
        Exit Function
    End If
    mess = mess & IIf(Len(mess), vbCrLf, "") & Replace(Replace(m_ValidateMsg, "<value>", NewValue), "<datatype>", VBVarType_ToStr(VbVarType.vbCurrency))
    If MsgBox(mess, vbOKCancel) = VbMsgBoxResult.vbOK Then
        'give back the users faulty new value, so user gets the chance to correct it
        Currency_TryParseValidate = NewValue
        Exit Function
    End If
    'give back the valid  old value, at least it was a valid Currency value
    Currency_TryParseValidate = IIf(Len(sFormat), Format(OldValueIn_NewValueOut, sFormat), OldValueIn_NewValueOut)
End Function

Public Function Decimal_TryParse(ByVal Value As String, ByRef Value_out) As Boolean
Try: On Error GoTo Catch
    If Len(DecimalSeparator) = 0 Then DecimalSeparator = Mid(CStr(0.1), 2, 1)
    Value = Replace(Value, ",", DecimalSeparator)
    Value = Replace(Value, ".", DecimalSeparator)
    Value_out = CDec(Value)
    Decimal_TryParse = True
    Exit Function
Catch:
End Function

Public Function Decimal_TryParseMess(ByVal Value As String, ByVal mess As String, ByRef Value_inout) As Boolean
    Decimal_TryParseMess = Decimal_TryParse(Value, Value_inout)
    If Not Decimal_TryParseMess Then MsgBox Replace(Replace(mess, "<value>", Value), "<datatype>", VBVarType_ToStr(VbVarType.vbDecimal))
End Function

Public Function String_TryParse(ByVal s As String, Value_out) As Boolean
Try: On Error GoTo Catch
    If Left(s, 1) = """" And Right(s, 1) = """" Then
        Value_out = s
        String_TryParse = True
    End If
Catch:
End Function

Public Function Identifier_TryParse(ByVal s As String, Value_out As String) As Boolean
Try: On Error GoTo Catch
    s = Trim(s)
    If Left(s, 1) = """" Then Exit Function
    If Right(s, 1) = """" Then Exit Function
    Dim i As Long: i = 1
    Select Case AscW(Mid(s, i, 1))
    Case 65 To 90, 95, 97 To 122 '"A" - "Z", "_", "a" - "z"
    Case Else: Exit Function
    End Select
    For i = 2 To Len(s)
        Select Case AscW(Mid(s, i, 1))
        Case 48 To 57, 65 To 90, 95, 97 To 122 '"A" - "Z", "_", "a" - "z"
        Case Else: Exit Function
        End Select
    Next
    Value_out = s
    Identifier_TryParse = True
Catch:
End Function

Public Function Array_TryParse(ByVal s As String, Value_out, Optional ByVal Delimiter = vbTab) As Boolean
Try: On Error GoTo Catch
    Value_out = Split(s, Delimiter)
    Array_TryParse = True
    Exit Function
Catch:
End Function

Public Function Array_ToStr(arr) As String
    Dim i As Long, s As String: s = "("
    If IsObject(arr(0)) Then
        For i = LBound(arr) To UBound(arr)
            s = s & arr(i).ToStr & "; "
        Next
    Else
        For i = LBound(arr) To UBound(arr)
            s = s & CStr(arr(i)) & "; "
        Next
    End If
    Array_ToStr = s & ")"
End Function

'
'there can either be
'&H8000  which is Integer = -32768
'or
'&H8000& which is Long    =  32768
'
'there can either be
'&HFFFF  which is Integer = -1
'or
'&HFFFF& which is Long    =  65535
'
'Private Const iiii  As Long    = &H8000&
'Private Const iiii2 As Integer = &H1234

'Public Function TestHexTryParse(ByVal s As String) As Boolean
'    Dim vt As VbVarType, v
'    If Hex_TryParse(s, vt, v) Then
'        Debug.Print s & " as " & VBVarType_ToStr(vt) & " = " & v
'        TestHexTryParse = True
'    End If
'End Function

Public Function HexInt_TryParse(ByVal s As String, ByRef Value_out As Integer) As Boolean
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    If IsHexPrefix(s) Then s = Mid(s, 3, l - 2): l = Len(s)
    Dim vt As VbVarType
    If VBTypeIdentifier_TryParse(s, vt) Then
        s = Left(s, l - 1)
        If vt <> VbVarType.vbInteger Then Exit Function
    End If
    If Not IsHex(s) Then Exit Function
    Value_out = CInt(s)
    HexInt_TryParse = True
Catch:
End Function

Public Function HexLng_TryParse(ByVal s As String, ByRef Value_out As Long) As Boolean
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    If IsHexPrefix(s) Then s = Mid(s, 3, l - 2): l = Len(s)
    Dim vt As VbVarType
    If VBTypeIdentifier_TryParse(s, vt) Then
        s = Left(s, l - 1)
        If vt <> VbVarType.vbLong Then Exit Function
    End If
    If Not IsHex(s) Then Exit Function
    Value_out = CLng(s)
    HexLng_TryParse = True
Catch:
End Function

Public Function Hex_TryParse(ByVal s As String, vtid_out As VbVarType, ByRef Value_out) As Boolean
    'string must be like
    '&H0, &H12, &HABCDEF12, &H12&,
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    '&HABCDEF12&
    '12345678901
    '
    If l < 3 Or 11 < l Then Exit Function
    If Not IsHexPrefix(s) Then Exit Function
    'there is either % for integer, or & for long or none
    If VBTypeIdentifier_TryParse(s, vtid_out) Then s = Left(s, l - 1)
    If Not IsHex(Mid(s, 3)) Then Exit Function
    Dim i As Integer, lng As Long
    If vtid_out = VbVarType.vbEmpty Then
        If l < 6 Then
            'vtid_out
            Value_out = CInt(s)
        Else
            Value_out = CLng(s)
        End If
    ElseIf vtid_out = VbVarType.vbLong Then
        Value_out = CLng(s)
    End If
    
    'Soll man den vbtypeid zurückgeben, oder nicht?
    'NEIN es soll nur vbtypeid zurückgegeben werden wenn einer hinten dranhängt
    'wenn kein vbtypeid hinten dranhängt dann soll empty zurückgegeben werden!
    
    vtid_out = vtid_out Or vbHex
    Hex_TryParse = True
Catch:
End Function

Public Function Hex_ToStr(ByVal Value) As String
    Dim s As String: s = "&H"
    Dim vt0 As VbVarType: vt0 = VarType(Value)
    Select Case vt0
    Case VbVarType.vbByte:     s = s & Hex(CByte(Value)) ' Hex2(CByte(Value))
    Case VbVarType.vbInteger:  s = s & Hex(CInt(Value))  ' Hex4(CInt(Value))
    Case VbVarType.vbLong:     s = s & Hex(CLng(Value))  ' Hex8(CLng(Value))
    Case VbVarType.vbCurrency: s = s & Hex16(CCur(Value))
    Case VbVarType.vbDecimal:  's = s & Hex32(CDec(Value)) ' ???
                               s = s & Hex16(CCur(Value))
    Case Else
        If vt0 And VbVarType.vbArray = VbVarType.vbArray Then
            Dim vt1 As VbVarType: vt1 = vt0 Xor VbVarType.vbArray
            Select Case vt1
            Case VbVarType.vbByte: s = s & ByteArray_ToHex(Value)
            End Select
        End If
    End Select
    Hex_ToStr = s
End Function

'Private Function ByteArray_ToHex(Bytes() As Byte) As String
Private Function ByteArray_ToHex(Value) As String
    Dim Bytes() As Byte: Bytes = Value
    Dim i As Long, s As String
    Dim lb As Long: lb = LBound(Bytes)
    Dim ub As Long: ub = UBound(Bytes)
    Dim n As Long: n = ub - lb + 1
    If n > 1024 Then Exit Function
    For i = lb To ub
        s = s & Hex2(Bytes(i))
    Next
    ByteArray_ToHex = s
End Function

Public Function OctInt_TryParse(ByVal s As String, ByRef Value_out As Integer) As Boolean
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    If IsOctPrefix(s) Then s = Mid(s, 3, l - 2): l = Len(s)
    Dim vt As VbVarType
    If VBTypeIdentifier_TryParse(s, vt) Then
        s = Left(s, l - 1)
        If vt <> VbVarType.vbInteger Then Exit Function
    End If
    If Not IsOct(s) Then Exit Function
    Value_out = CInt(s)
    OctInt_TryParse = True
Catch:
End Function

Public Function OctLng_TryParse(ByVal s As String, ByRef Value_out As Long) As Boolean
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    If IsOctPrefix(s) Then s = Mid(s, 3, l - 2): l = Len(s)
    Dim vt As VbVarType
    If VBTypeIdentifier_TryParse(s, vt) Then
        s = Left(s, l - 1)
        If vt <> VbVarType.vbLong Then Exit Function
    End If
    If Not IsOct(s) Then Exit Function
    Value_out = CLng(s)
    OctLng_TryParse = True
Catch:
End Function

Public Function Oct_TryParse(ByVal s As String, vtid_out As VbVarType, Value_out) As Boolean
    'string must be like
    '&H0, &H12, &H12345670, &H12&,
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    '&H12345678&
    '12345678901
    '
    If l < 3 Or 11 < l Then Exit Function
    If Not IsOctPrefix(s) Then Exit Function
    'there is either % for integer, or & for long or none
    If VBTypeIdentifier_TryParse(s, vtid_out) Then s = Left(s, l - 1)
    If Not IsOct(Mid(s, 3)) Then Exit Function
    If vtid_out = VbVarType.vbEmpty Then
        If l < 6 Then
            Value_out = CInt(s)
        Else
            Value_out = CLng(s)
        End If
    ElseIf vtid_out = VbVarType.vbLong Then
        Value_out = CLng(s)
    End If
    Oct_TryParse = True
Catch:
End Function

Public Function Oct_ToStr(ByVal Value) As String
    Dim s As String: s = "&O"
    Select Case VarType(Value)
    Case VbVarType.vbByte:     s = s & Oct3(CByte(Value))
    Case VbVarType.vbInteger:  s = s & Oct6(CInt(Value))
    Case VbVarType.vbLong:     s = s & Oct11(CLng(Value))
    Case VbVarType.vbCurrency: s = s & Oct22(CCur(Value))
    'Case VbVarType.vbDecimal:  s = s & Hex32(CDec(Value))
    'Case VbVarType.vbDecimal:  s = s & Hex16(CCur(Value))
    End Select
    Oct_ToStr = s
End Function

Public Function BinInt_TryParse(ByVal s As String, ByRef Value_out As Integer) As Boolean
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    If IsBinPrefix(s) Then s = Mid(s, 3, l - 2): l = Len(s)
    Dim vt As VbVarType
    If VBTypeIdentifier_TryParse(s, vt) Then
        s = Left(s, l - 1): l = Len(s)
        If vt <> VbVarType.vbInteger Then Exit Function
    End If
    If Not IsBin(s) Then Exit Function
    If 16 < l Then Exit Function
    Dim i As Long, n As Long: n = MMath.Min(l, 15)
    Dim v As Integer
    For i = 0 To n - 1
        If Mid(s, l - i, 1) = "1" Then v = v + 2 ^ i
    Next
    If l = 16 Then
        If Mid(s, l - i, 1) = "1" Then v = v Xor &H8000
    End If
    Value_out = v
    BinInt_TryParse = True
Catch:
End Function

Public Function BinInt_ToStr(ByVal Value As Integer) As String
    'with or without starting 0 ?
    'here for now first with all starting 0
    Dim s As String
    Dim i As Long, v As Integer
    For i = 0 To 14
        v = 2 ^ i
        If Value And v Then
            s = "1" & s
        Else
            s = "0" & s
        End If
    Next
    If Value < 0 Then s = "1" & s Else s = "0" & s
    BinInt_ToStr = s
End Function

Public Function BinLng_TryParse(ByVal s As String, ByRef Value_out As Long) As Boolean
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    If IsBinPrefix(s) Then s = Mid(s, 3, l - 2): l = Len(s)
    Dim vt As VbVarType
    If VBTypeIdentifier_TryParse(s, vt) Then
        s = Left(s, l - 1): l = Len(s)
        If vt <> VbVarType.vbLong Then Exit Function
    End If
    If Not IsBin(s) Then Exit Function
    If 32 < l Then Exit Function
    Dim i As Long, n As Long: n = Min(l, 31)
    Dim v As Long
    For i = 0 To n - 1
        If Mid(s, l - i, 1) = "1" Then v = v + 2 ^ i
    Next
    If l = 32 Then
        If Mid(s, l - i, 1) = "1" Then v = v Xor &H80000000
    End If
    Value_out = v
    BinLng_TryParse = True
Catch:
End Function

Public Function BinLng_ToStr(ByVal Value As Long) As String
    'with or without starting 0 ?
    'here first with starting 0
    Dim s As String
    Dim i As Long, v As Long
    For i = 0 To 30
        v = 2 ^ i
        If Value And v Then
            s = "1" & s
        Else
            s = "0" & s
        End If
    Next
    If Value < 0 Then s = "1" & s Else s = "0" & s
    BinLng_ToStr = s
End Function

Public Function Bin_TryParse(ByVal s As String, vtid_out As VbVarType, Value_out) As Boolean
    'string must be like
    '&B0, &B10, &B1010100110, &B10&,
Try: On Error GoTo Catch
    s = Trim(s)
    Dim l As Long: l = Len(s)
    If l < 3 Or 11 < l Then Exit Function
    If Not IsBinPrefix(s) Then Exit Function
    'there is either % for integer, or & for long or none
    If VBTypeIdentifier_TryParse(s, vtid_out) Then s = Left(s, l - 1)
    If Not IsBin(Mid(s, 3)) Then Exit Function
    Dim i As Integer, lng As Long
    Bin_TryParse = BinInt_TryParse(s, i)
    If Bin_TryParse Then
        vtid_out = vtid_out Or vbBin
        Value_out = i
        Exit Function
    End If
    Bin_TryParse = BinLng_TryParse(s, lng)
    If Bin_TryParse Then
        vtid_out = vtid_out Or vbBin
        Value_out = l
        Exit Function
    End If
Catch:
End Function

Public Function Bin_ToStr(ByVal Value) As String
    Dim s As String: s = "&B"
    Select Case VarType(Value)
    Case VbVarType.vbByte:     s = s & Bin8(CByte(Value))
    Case VbVarType.vbInteger:  s = s & Bin16(CInt(Value))
    Case VbVarType.vbLong:     s = s & Bin32(CLng(Value))
    Case VbVarType.vbCurrency: s = s & Bin64(CCur(Value))
    Case VbVarType.vbDecimal:  's = s & Bin96(CDec(Value))
                               's = s & Bin96(CCur(Value))
    End Select
    Bin_ToStr = s
End Function

Public Function CheckType(ByVal s As String, ByVal vt As VbVarType, Value_out) As Boolean
    'checks if the type vt is inside s and returns the value in value_out
    If vt And vbHex Then vt = vt Xor vbHex
    If vt And vbOct Then vt = vt Xor vbOct
    If vt And vbBin Then vt = vt Xor vbBin

    Select Case vt
    Case VbVarType.vbByte:        Dim BytVal As Byte:     CheckType = Byte_TryParse(s, BytVal):      Value_out = BytVal
    Case VbVarType.vbInteger:     Dim IntVal As Integer:  CheckType = Integer_TryParse(s, IntVal):   Value_out = IntVal
    Case VbVarType.vbBoolean:     Dim BolVal As Boolean:  CheckType = Boolean_TryParse(s, BolVal):   Value_out = BolVal
    Case VbVarType.vbLong:        Dim LngVal As Long:     CheckType = Long_TryParse(s, LngVal):      Value_out = LngVal
    Case VbVarType.vbSingle:      Dim SngVal As Single:   CheckType = Single_TryParse(s, SngVal):    Value_out = SngVal
    Case VbVarType.vbDouble:      Dim DblVal As Double:   CheckType = Double_TryParse(s, DblVal):    Value_out = DblVal
    Case VbVarType.vbCurrency:    Dim CurVal As Currency: CheckType = Currency_TryParse(s, CurVal):  Value_out = CurVal
    Case VbVarType.vbDecimal:     Dim DecVal As Variant:  CheckType = Decimal_TryParse(s, DecVal):   Value_out = DecVal
    Case VbVarType.vbDate:        Dim DatVal As Date:     CheckType = Date_TryParse(s, DatVal):      Value_out = DatVal
    Case VbVarType.vbString:      Dim StrVal As String:   CheckType = String_TryParse(s, StrVal):    Value_out = StrVal
    Case VbVarType.vbArray:       Dim ArrVal As Variant:  CheckType = Array_TryParse(s, ArrVal):     Value_out = ArrVal
    End Select
End Function

Public Function VBVarType_TryParse(ByVal s As String, vt_out As VbVarType) As Boolean
    'returns true if s matches any vb-datatype notations and returns the datatype in vt_out
    s = UCase(Trim(s))
    Select Case s
    Case "INTEGER":         vt_out = VbVarType.vbInteger
    Case "LONG":            vt_out = VbVarType.vbLong
    Case "SINGLE":          vt_out = VbVarType.vbSingle
    Case "DOUBLE":          vt_out = VbVarType.vbDouble
    Case "CURRENCY":        vt_out = VbVarType.vbCurrency
    Case "DATE":            vt_out = VbVarType.vbDate
    Case "STRING":          vt_out = VbVarType.vbString
    Case "OBJECT":          vt_out = VbVarType.vbObject
    Case "ERR":             vt_out = VbVarType.vbError
    Case "ERROR":           vt_out = VbVarType.vbError
    Case "BOOLEAN":         vt_out = VbVarType.vbBoolean
    Case "VARIANT":         vt_out = VbVarType.vbVariant
    Case "DATAOBJECT":      vt_out = VbVarType.vbDataObject
    Case "DECIMAL":         vt_out = VbVarType.vbDecimal
    Case "BYTE":            vt_out = VbVarType.vbByte
    Case "USERDEFINEDTYPE": vt_out = VbVarType.vbUserDefinedType
    Case "ARRAY":           vt_out = VbVarType.vbArray
    Case Else: Exit Function
    End Select
    VBVarType_TryParse = True
End Function

Public Function VBVarType_ToStr(ByVal vt As VbVarType) As String
    'returns a string representation of the vb-datatype in vt
    'also have a look in MVariant
    Dim s As String
    
    If vt And vbHex Then s = " (Hex)": vt = vt Xor vbHex
    If vt And vbOct Then s = " (Oct)": vt = vt Xor vbOct
    If vt And vbBin Then s = " (Bin)": vt = vt Xor vbBin
    
    Select Case vt
    Case VbVarType.vbEmpty:           s = "None/Empty" & s ' =  0
    Case VbVarType.vbNull:            s = "Null" & s        ' =  1
    Case VbVarType.vbInteger:         s = "Integer" & s     ' =  2
    Case VbVarType.vbLong:            s = "Long" & s        ' =  3
    Case VbVarType.vbSingle:          s = "Single" & s      ' =  4
    Case VbVarType.vbDouble:          s = "Double" & s      ' =  5
    Case VbVarType.vbCurrency:        s = "Currency" & s    ' =  6
    Case VbVarType.vbDate:            s = "Date" & s        ' =  7
    Case VbVarType.vbString:          s = "String" & s      ' =  8
    Case VbVarType.vbObject:          s = "Object" & s      ' =  9
    Case VbVarType.vbError:           s = "Error" & s       ' = 10
    Case VbVarType.vbBoolean:         s = "Boolean" & s     ' = 11
    Case VbVarType.vbVariant:         s = "Variant" & s     ' = 12
    Case VbVarType.vbDataObject:      s = "DataObject" & s  ' = 13
    Case VbVarType.vbDecimal:         s = "Decimal" & s     ' = 14
    Case VbVarType.vbByte:            s = "Byte" & s       ' = 17 (&H11)
    Case VbVarType.vbUserDefinedType: s = "UserDefinedType" & s ' = 36 (&H24)
    Case VbVarType.vbArray:           s = "Array" & s      ' = 8192 (&H2000)
    
    End Select
    VBVarType_ToStr = s
End Function

Public Function VBVarType_IsNumeric(ByVal vt As VbVarType) As Boolean
    Select Case True
    Case vt And VbVarType.vbByte:     VBVarType_IsNumeric = True
    Case vt And VbVarType.vbInteger:  VBVarType_IsNumeric = True
    Case vt And VbVarType.vbLong:     VBVarType_IsNumeric = True
    Case vt And VbVarType.vbSingle:   VBVarType_IsNumeric = True
    Case vt And VbVarType.vbDouble:   VBVarType_IsNumeric = True
    Case vt And VbVarType.vbCurrency: VBVarType_IsNumeric = True
    Case vt And VbVarType.vbDecimal:  VBVarType_IsNumeric = True
    End Select
End Function

'vbEmpty           =  0
'vbNull            =  1
'vbInteger         =  2
'vbLong            =  3
'vbSingle          =  4
'vbDouble          =  5
'vbCurrency        =  6
'vbDate            =  7
'vbString          =  8
'vbObject          =  9
'vbError           = 10
'vbBoolean         = 11
'vbVariant         = 12
'vbDataObject      = 13
'vbDecimal         = 14
'vbByte            = 17 (&H11)
'vbUserDefinedType = 36 (&H24)
'vbArray           = 8192 (&H2000)
'vbHex             = &H10000
'vbOct             = &H20000
'vbBin             = &H40000

'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary
Public Function VBTypeIdentifier_TryParse(ByVal s As String, vt_out As VbVarType) As Boolean
    'returns true if s has any vb-type-identifier attached to it and returns the vbtypeid in vt_out
'%   Ganze Zahl  Dim L%
'&   Long        Dim M&
'^   LongLong    Dim N^
'@   Währung     Const W@ = 37.5
'!   Single      Dim Q!
'#   Gleitkommawert mit doppelter Genauigkeit
'                Dim X#
'$   String      Dim V$ = "Secret"
    
    'returns true if s has a postfix-type-identifier,
    'the corresponding datatype will be returned in vt_out
    'in a second step check if the value matches the type identifier
    
    If Len(s) = 0 Then Exit Function
    Dim b As Boolean, ti As String: ti = Right(s, 1)
    Select Case ti
    Case "%": vt_out = VbVarType.vbInteger:  b = True
    Case "&": vt_out = VbVarType.vbLong:     b = True
#If VBA7 Then
    Case "^": vt_out = VbVarType.vbLongLong: b = True
#End If
    Case "@": vt_out = VbVarType.vbCurrency: b = True
    Case "!": vt_out = VbVarType.vbSingle:   b = True
    Case "#": vt_out = VbVarType.vbDouble:   b = True
    Case "$": vt_out = VbVarType.vbString:   b = True
    End Select
    VBTypeIdentifier_TryParse = b
End Function

Public Function VBTypeIdentifier_ToStr(ByVal vtid As VbVarType) As String
    Dim s As String
    Select Case vtid
    Case VbVarType.vbInteger:  s = "%"
    Case VbVarType.vbLong:     s = "&"
#If VBA7 Then
    Case VbVarType.vbLongLong: s = "^"
#End If
    Case VbVarType.vbCurrency: s = "@"
    Case VbVarType.vbSingle:   s = "!"
    Case VbVarType.vbDouble:   s = "#"
    Case VbVarType.vbString:   s = "$"
    End Select
    VBTypeIdentifier_ToStr = s
End Function

Public Function Numeric_TryParse(ByVal s As String, vtid_out As Long, Value_out As Variant) As Boolean
    'returns true if s contains any numeric datatype and returns the optional vbtype-id in vtid_out and the value in value_out
Try: On Error GoTo Catch
    
    s = Trim(s): If Len(s) = 0 Then Exit Function
    
    Numeric_TryParse = Hex_TryParse(s, vtid_out, Value_out)
    If Numeric_TryParse Then vtid_out = vtid_out Or vbHex: Exit Function
    
    Numeric_TryParse = Oct_TryParse(s, vtid_out, Value_out)
    If Numeric_TryParse Then vtid_out = vtid_out Or vbOct: Exit Function
    
    Numeric_TryParse = Bin_TryParse(s, vtid_out, Value_out)
    If Numeric_TryParse Then vtid_out = vtid_out Or vbBin: Exit Function
    
    vtid_out = 0
    If VBTypeIdentifier_TryParse(s, vtid_out) Then
        s = Left(s, Len(s) - 1)
    End If
    
    Dim byt As Byte:    Numeric_TryParse = Byte_TryParse(s, byt)
    If Numeric_TryParse Then Value_out = byt: Exit Function
    
    Dim iii As Integer: Numeric_TryParse = Integer_TryParse(s, iii)
    If Numeric_TryParse Then Value_out = iii: Exit Function
    
    Dim lng As Long:    Numeric_TryParse = Long_TryParse(s, lng)
    If Numeric_TryParse Then Value_out = lng: Exit Function
    
    'what distinguishes between Single or Double?
    'the length of the string in characters?
    'nope everything what has a period is a Double
    'everything what has a type-identifier is of this type identifier
    
    Select Case vtid_out
    Case VbVarType.vbSingle
        Dim sng As Single:   Numeric_TryParse = Single_TryParse(s, sng)
        If Numeric_TryParse Then Value_out = sng: Exit Function
    Case VbVarType.vbCurrency
        Dim cur As Currency: Numeric_TryParse = Currency_TryParse(s, cur)
        If Numeric_TryParse Then Value_out = cur: Exit Function
    Case VbVarType.vbDecimal
        Dim dec:             Numeric_TryParse = Decimal_TryParse(s, cur)
        If Numeric_TryParse Then Value_out = dec: Exit Function
    Case Else
        Dim dbl As Double:   Numeric_TryParse = Double_TryParse(s, dbl)
        If Numeric_TryParse Then Value_out = dbl: Exit Function
    End Select
    
Catch:
End Function

'    Dim l As Long: l = Len(s)
'    If l = 0 Then Exit Function
'    'returns true if s contains a numeric value
'    'this could be: Byte, Integer, Long, Single, Double, Currency, Decimal
'    'for int-types this could be Hex, Oct, Bin or Decimal
'    'for flt-types this could be Single, Double, Currency oder Decimal
'    's could have a typeidentifier-character at the end
'    'then we must check whether the typeidentifier matches the type of the value
'    'in vtid_out we return the type-identifier and additional bits for hex, oct, bin
'    'hex =
'    Dim vtid As VbVarType
'    If VBTypeIdentifier_TryParse(s, vtid) Then
'        s = Left(s, Len(s) - 1)
'        l = l - 1
'        If l = 0 Then Exit Function
'    End If
'    Dim bHex As Boolean: bHex = IsHexPrefix(s): If bHex Then vtid = vtid Or &H10000
'    Dim bOct As Boolean: bOct = IsOctPrefix(s): If bOct Then vtid = vtid Or &H20000
'    Dim bBin As Boolean: bBin = IsBinPrefix(s): If bBin Then vtid = vtid Or &H40000
'    If bHex Or bOct Or bBin Then s = Mid(s, 3)
'    Dim lng As Long
'    If bHex And IsHex(s) Then
'        l = Len(s)
'        lng = CLng("&H" & s)
'        Select Case True
'        Case Is <= 2: If vtid = VbVarType.vbEmpty Then Numeric_TryParse = True
'                Value_out = CByte(lng)
'        Case Is <= 4: If vtid = VbVarType.vbInteger Then Numeric_TryParse = True
'                Value_out = CInt(lng)
'        Case Is <= 8: If vtid = VbVarType.vbLong Then Numeric_TryParse = True
'                Value_out = lng
'        Case Else
'        End Select
'
'        Exit Function
'    End If
'
'
'        s = Mid(s, 3, Len(s) - 3)
'        If IsHex(s) Then
'            'v_out =
'            Exit Function
'        End If
'    End If

Public Function Literal_TryParse(ByVal s As String, vtid_out As Long, Value_out As Variant) As Boolean
    
    If Len(s) = 0 Then Exit Function
    
    vtid_out = 0
    
    Literal_TryParse = Numeric_TryParse(s, vtid_out, Value_out)
    If Literal_TryParse Then Exit Function
    
    Dim bol As Boolean:  Literal_TryParse = Boolean_TryParse(s, bol)
    If Literal_TryParse Then Value_out = bol: Exit Function
    
    Dim vbid As String:  Literal_TryParse = Identifier_TryParse(s, vbid)
    If Literal_TryParse Then Value_out = vbid: vtid_out = vtid_out Or vbIdent: Exit Function
    
    Dim str  As String:  Literal_TryParse = String_TryParse(s, str)
    If Literal_TryParse Then Value_out = str: Exit Function
    
    Dim dat  As Date:    Literal_TryParse = Date_TryParse(s, dat)
    If Literal_TryParse Then Value_out = dat: Exit Function
    
    'parsing if its an expression?
End Function

' ^ ############################## ^ '    TryParse Functions    ' ^ ############################## ^ '

' v ############################## v '    Hex, Dec, Oct, Bin ToStr Functions    ' v ############################## v '

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

Public Function Dec2(ByVal Value As Long) As String
    Dec2 = CStr(Value): If Len(Dec2) < 2 Then Dec2 = "0" & Dec2
End Function

Public Function Oct3(ByVal Value As Byte) As String
    Oct3 = Oct(Value): If Len(Oct3) < 3 Then Oct3 = String(3 - Len(Oct3), "0") & Oct3
End Function

Public Function Oct6(ByVal Value As Integer) As String
    Oct6 = Oct(Value): If Len(Oct6) < 6 Then Oct6 = String(6 - Len(Oct6), "0") & Oct6
End Function

Public Function Oct11(ByVal Value As Long) As String
    Oct11 = Oct(Value): If Len(Oct11) < 11 Then Oct11 = String(11 - Len(Oct11), "0") & Oct11
End Function

Public Function Oct22(ByVal Value As Currency) As String
    Dim tc As TCur:  tc.Value = Value
    Dim tl As TLong2: LSet tl = tc
    Oct22 = Oct11(tl.Hi) & Oct11(tl.Lo) 'no this does not work
End Function

Public Function Bin8(ByVal Value As Byte) As String
    'with or without starting 0 ?
    'here for now first with all starting 0
    Dim s As String
    Dim i As Long, v As Byte
    For i = 0 To 7
        v = 2 ^ i
        If Value And v Then
            s = "1" & s
        Else
            s = "0" & s
        End If
    Next
    'If Value < 0 Then s = "1" & s Else s = "0" & s
    Bin8 = s
End Function

Public Function Bin16(ByVal Value As Integer) As String
    Bin16 = BinInt_ToStr(Value)
End Function

Public Function Bin32(ByVal Value As Long) As String
    Bin32 = BinLng_ToStr(Value)
End Function

Public Function Bin64(ByVal Value As Currency) As String
    Dim tc As TCur:  tc.Value = Value
    Dim tl As TLong2: LSet tl = tc
    Bin64 = Bin32(tl.Hi) & Bin32(tl.Lo)
End Function

'ulong ul = ulong.MaxValue;
'long l = (long)ul;
'var s = Convert.ToString(l, 8); //8 => oct, 2 => bin
'Console.WriteLine(s); //Outputs 1777777777777777777777

'byte.MaxValue   = (Bin)11111111, (Oct)377, (Dec)255, (Hex)FF
'byte.MinValue   = (Bin)00000000, (Oct)000, (Dec)000, (Hex)00

'sbyte.MaxValue  = (Bin)11111111, (Oct)377, (Dec)255, (Hex)FF
'sbyte.MinValue  = (Bin)00000000, (Oct)000, (Dec)000, (Hex)00

'short.MaxValue  = (Bin)0111111111111111, (Oct)077777, (Dec)32767, (Hex)7FFF
'short.-1        = (Bin)0111111111111111, (Oct)177777, (Dec)-00001, (Hex)FFFF

'ushort.MaxValue = (Bin)1111111111111111, (Oct)177777, (Dec)65535, (Hex)FFFF
'ushort.MinValue = (Bin)0000000000000000, (Oct)000000, (Dec)00000, (Hex)0000

'int.MaxValue    = (Bin)01111111111111111111111111111111, (Oct)17777777777, (Dec) 2147483647, (Hex)7FFFFFFF
'int.-1          = (Bin)11111111111111111111111111111111, (Oct)37777777777, (Dec)-0000000001, (Hex)FFFFFFFF

'uint.MaxValue   = (Bin)11111111111111111111111111111111, (Oct)37777777777, (Dec)4294967295, (Hex)FFFFFFFF
'uint.MinValue   = (Bin)00000000000000000000000000000000, (Oct)00000000000, (Dec)0000000000, (Hex)00000000

'long.MaxValue   = (Bin)0111111111111111111111111111111111111111111111111111111111111111, (Oct)0777777777777777777777, (Dec) 9223372036854775807,  (Hex)7FFFFFFFFFFFFFFF
'long.-1         = (Bin)1111111111111111111111111111111111111111111111111111111111111111, (Oct)1777777777777777777777, (Dec)-00000000000000000001, (Hex)FFFFFFFFFFFFFFFF

'ulong.MaxValue  = (Bin)1111111111111111111111111111111111111111111111111111111111111111, (Oct)1777777777777777777777, (Dec)-00000000000000000001, (Hex)FFFFFFFFFFFFFFFF
'ulong.MinValue  = (Bin)0000000000000000000000000000000000000000000000000000000000000000, (Oct)0000000000000000000000, (Dec) 00000000000000000000, (Hex)0000000000000000

'BasedValueConverter bvc;
'bvc = new BasedValueConverter(byte.MaxValue.Dump("BasedValueConverter byte"));
'bvc.ToStringBin().Dump();  // 01111111
'bvc.ToStringOct().Dump();  // 177
'bvc.ToStringDec().Dump();  // 127
'bvc.ToStringHex().Dump();  // 7F
'bvc = new BasedValueConverter(byte.MinValue);
'bvc.ToStringBin().Dump();  // 00000000
'bvc.ToStringOct().Dump();  // 000
'bvc.ToStringDec().Dump();  // 000
'bvc.ToStringHex().Dump();  // 00
'
'bvc = new BasedValueConverter(sbyte.MaxValue.Dump("BasedValueConverter sbyte"));
'bvc.ToStringBin().Dump();  // 01111111
'bvc.ToStringOct().Dump();  // 177
'bvc.ToStringDec().Dump();  // 127
'bvc.ToStringHex().Dump();  // 7F
'bvc = new BasedValueConverter((sbyte)-1);
'bvc.ToStringBin().Dump();  // 11111111
'bvc.ToStringOct().Dump();  // 377
'bvc.ToStringDec().Dump();  // -001
'bvc.ToStringHex().Dump();  // FF
'
'bvc = new BasedValueConverter(short.MaxValue.Dump("BasedValueConverter short"));
'bvc.ToStringBin().Dump();  // 0111111111111111
'bvc.ToStringOct().Dump();  // 077777
'bvc.ToStringDec().Dump();  // 32767
'bvc.ToStringHex().Dump();  // 7FFF
'bvc = new BasedValueConverter((short)-1);
'bvc.ToStringBin().Dump();  // 1111111111111111
'bvc.ToStringOct().Dump();  // 177777
'bvc.ToStringDec().Dump();  // -00001
'bvc.ToStringHex().Dump();  // FFFF
'
'bvc = new BasedValueConverter(ushort.MaxValue.Dump("BasedValueConverter ushort"));
'bvc.ToStringBin().Dump();  // 1111111111111111
'bvc.ToStringOct().Dump();  // 177777
'bvc.ToStringDec().Dump();  // 65535
'bvc.ToStringHex().Dump();  // FFFF
'bvc = new BasedValueConverter(ushort.MinValue);
'bvc.ToStringBin().Dump();  // 0000000000000000
'bvc.ToStringOct().Dump();  // 000000
'bvc.ToStringDec().Dump();  // 00000
'bvc.ToStringHex().Dump();  // 0000
'
'bvc = new BasedValueConverter(int.MaxValue.Dump("BasedValueConverter int"));
'bvc.ToStringBin().Dump();  // 01111111111111111111111111111111
'bvc.ToStringOct().Dump();  // 17777777777
'bvc.ToStringDec().Dump();  // 2147483647
'bvc.ToStringHex().Dump();  // 7FFFFFFF
'bvc = new BasedValueConverter((int)-1);
'bvc.ToStringBin().Dump();  // 11111111111111111111111111111111
'bvc.ToStringOct().Dump();  // 37777777777
'bvc.ToStringDec().Dump();  // -0000000001
'bvc.ToStringHex().Dump();  // FFFFFFFF
'
'bvc = new BasedValueConverter(uint.MaxValue.Dump("BasedValueConverter uint"));
'bvc.ToStringBin().Dump();  // 11111111111111111111111111111111
'bvc.ToStringOct().Dump();  // 37777777777
'bvc.ToStringDec().Dump();  // 4294967295
'bvc.ToStringHex().Dump();  // FFFFFFFF
'bvc = new BasedValueConverter(uint.MinValue);
'bvc.ToStringBin().Dump();  // 00000000000000000000000000000000
'bvc.ToStringOct().Dump();  // 00000000000
'bvc.ToStringDec().Dump();  // 0000000000
'bvc.ToStringHex().Dump();  // 00000000
'
'bvc = new BasedValueConverter(long.MaxValue.Dump("BasedValueConverter long"));
'bvc.ToStringBin().Dump();  // 0111111111111111111111111111111111111111111111111111111111111111
'bvc.ToStringOct().Dump();  // 0777777777777777777777
'bvc.ToStringDec().Dump();  // 9223372036854775807
'bvc.ToStringHex().Dump();  // 7FFFFFFFFFFFFFFF
'bvc = new BasedValueConverter((long)-1);
'bvc.ToStringBin().Dump();  // 1111111111111111111111111111111111111111111111111111111111111111
'bvc.ToStringOct().Dump();  // 1777777777777777777777
'bvc.ToStringDec().Dump();  // -00000000000000000001
'bvc.ToStringHex().Dump();  // FFFFFFFFFFFFFFFF
'
'
'bvc = new BasedValueConverter(ulong.MaxValue.Dump("BasedValueConverter ulong"));
'bvc.ToStringBin().Dump();  // 1111111111111111111111111111111111111111111111111111111111111111
'bvc.ToStringOct().Dump();  // 1777777777777777777777
'bvc.ToStringDec().Dump();  // -00000000000000000001
'bvc.ToStringHex().Dump();  // FFFFFFFFFFFFFFFF
'bvc = new BasedValueConverter(ulong.MinValue);
'bvc.ToStringBin().Dump();  // 0000000000000000000000000000000000000000000000000000000000000000
'bvc.ToStringOct().Dump();  // 0000000000000000000000
'bvc.ToStringDec().Dump();  // 00000000000000000000
'bvc.ToStringHex().Dump();  // 0000000000000000

' ^ ############################## ^ '    Hex, Oct, Bin ToStr Functions    ' ^ ############################## ^ '

' v ' ############################## ' v '    System.String functions    ' v ' ############################## ' v '

Public Function Contains(s As String, ByVal Value As String) As Boolean
    Contains = InStr(1, s, Value) > 0
End Function

Public Function ContainsOneOf(s As String, Values() As String) As Boolean
    'returns true if s contains minimum one of the Values
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    For i = LBound(Values) To UBound(Values)
        ContainsOneOf = InStr(1, s, Values(i)) > 0
        If ContainsOneOf Then Exit Function
    Next
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

Function CHexToVBHex(ByVal s As String) As String
    CHexToVBHex = s
    If Left(s, 2) = "0x" Then CHexToVBHex = "&H" & Mid$(s, 3)
End Function

Public Function PadCentered(this As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String) As String
    Dim StringLength As Long: StringLength = Len(this)
    If StringLength > totalWidth Then
        PadCentered = this
    Else
        Dim l As Long: l = (totalWidth - StringLength) \ 2
        Dim r As Long: r = (totalWidth - StringLength) / 2
        If Len(paddingChar) Then
            PadCentered = String$(l, paddingChar) & this & String$(r, paddingChar)
        Else
            PadCentered = Space$(totalWidth)
            RSet PadCentered = this & Space$(r)
        End If
    End If
End Function

Function PadRight(this As String, ByVal totalWidth As Long, Optional ByVal paddingChar As String) As String
    
    ' Returns the String this with the specified length.
    ' totalWidth: the length of the returned string
    '             if totalWidth is smaller then the length of this then
    '             this will be returned
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
    Dim l As Long: l = Len(s)
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
        If startIndex < l Then
            Remove = Left$(s, startIndex)
            Exit Function
        End If
        If startIndex = l Then
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
        If startIndex < l Then
            Remove = s
            Exit Function
        End If
        If startIndex = l Then
            Remove = s
            Exit Function
        End If
        Remove = s
        'Error message
        Exit Function
    End If
    If Count < l Then
        If startIndex < 0 Then
            Remove = ""
            'Error message
            Exit Function
        End If
        If startIndex = 0 Then
            Remove = Mid$(s, Count + 1)
            Exit Function
        End If
        If startIndex < l Then
            If startIndex + Count < l Then
                Remove = Left(s, startIndex) & Mid(s, startIndex + Count + 1)
                Exit Function
            End If
            If startIndex + Count = l Then
                Remove = Left(s, startIndex)
                Exit Function
            End If
            If l < startIndex + Count Then
                Remove = Left(s, startIndex)
                'Error message
                Exit Function
            End If
        End If
        If startIndex = l Then
            Remove = s
            'Error message
            Exit Function
        End If
        Remove = s
        'Error message
        Exit Function
    End If
    If Count = l Then
        If startIndex < 0 Then
            Remove = ""
            'Error message
            Exit Function
        End If
        If startIndex = 0 Then
            Remove = "" 'Mid$(s, Count + 1)
            Exit Function
        End If
        If startIndex < l Then
            Remove = Left$(s, startIndex)
            'Error message
            Exit Function
        End If
    End If
    If l < Count Then
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
        If startIndex < l Then
            Remove = Left(s, startIndex)
            'Error message
            Exit Function
        End If
    End If
End Function

Public Function RemoveFromRightStartingWith(ByVal s As String, ByVal sRightStartingWith As String) As String
    'removes all characters from the starting of sRightStartingWith
    'and returns the remaining string
    Dim pos As Long: pos = InStr(1, s, sRightStartingWith)
    If pos <= 0 Then
        RemoveFromRightStartingWith = s
        Exit Function
    End If
    RemoveFromRightStartingWith = Left$(s, pos)
End Function

Public Function StartsWith(s As String, ByVal Value As String) As Boolean
    StartsWith = Left$(s, Len(Value)) = Value
End Function

Public Function Substring(s As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long) As String
    Substring = Mid(s, startIndex, Length)
End Function

Public Function Between(Line As String, ByVal s1 As String, ByVal s2 As String) As String
    'returns the string inside Line in between s1 and s2
    Dim pos1 As Long, pos2 As Long
    If Len(s1) Then
        pos1 = InStr(1, Line, s1)
        If pos1 <= 0 Then Exit Function
        pos1 = pos1 + Len(s1)
    Else
        pos1 = 1
    End If
    If Len(s2) Then
        pos2 = InStr(pos1, Line, s2)
        If pos2 <= 0 Then Exit Function
        pos2 = pos2 - 1 'Len(s)
    Else
        pos2 = Len(Line)
    End If
    Between = Mid(Line, pos1, pos2 - pos1 + 1)
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

Public Sub SCArray(sa() As String, ParamArray strArr())
    Dim i As Long: For i = 0 To UBound(strArr): sa(i) = strArr(i): Next
End Sub

Public Function AdverbNum_ToStr(ByVal num As Byte) As String
    Static sa(0 To 11) As String
    If Len(sa(1)) = 0 Then SCArray sa, "first", "second", "third", "fourth", "fifth", "sixt", "seventh", "eigth", "nineth", "tenth", "eleventh", "twelfth"
    AdverbNum_ToStr = sa(num - 1)
End Function
' ^ ' ############################## ' ^ '    System.String functions    ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    Unicode-BOM functions    ' v ' ############################## ' v '

'
'    Zeichenfolge:        4  3  2  1
'    bom_None       =              0
'    bom_UTF_32_BE  = &H FF FE 00 00  ' Big Endian Motorola '   0   0 254 255     ' ??þÿ
'    bom_SCSU       = &H    FF FE 0E  '                     '      14 254 255     ' ?þÿ      ' [6]
'    bom_UTF_7      = &H    76 2F 2B  '                     '      43  47 118     ' +/v
'    bom_GB_18030   = &H 33 95 31 84  '                     ' 132  49 149  51     ' „1•3
'    bom_UTF_EBCDIC = &H 73 66 73 DD  '                     ' 221 115 102 115     ' Ýsfs
'    bom_UTF_8      = &H    BF BB EF  '                     '     239 187 191     ' ï»¿     ' [4]
'    bom_UTF_1      = &H    4C 64 F7  '                     '     247 100  76     ' ÷dL
'    bom_BOCU_1     = &H    28 EE FB  '                     '     251 238  40     ' ûî(
'    bom_UTF_16_BE  = &H       FF FE  ' Big Endian Motorola '         254 255     ' þÿ
'    bom_UTF_16_LE  = &H       FE FF  ' little endian Intel '         255 254     ' ÿþ
'    bom_UTF_32_LE  = &H 00 00 FE FF  ' little endian Intel ' 255 254   0   0     ' ??ÿþ


'                                     ' und ein Zeichen aus: [ 56 | 57 | 43 | 47 ]
'                                     ' und ein Zeichen aus: [ 38 | 39 | 2B | 2F ]           ' [5]
'                                     ' +/v und ein Zeichen aus:  [  8 |  9 |  + |  / ]
'                                     ' (von anderen möglichen Bytefolgen wird abgeraten)
'                                     ' optional gefolgt von FF                              ' [7]
'                                     ' optional gefolgt von 255     ûî
'                                     ' optional gefolgt von          ÿ
Public Function IsBOM(ByVal s As String, Optional rest_out As String) As EByteOrderMark
    'checks if s starts with any BOM, returns the bom, andalso returns the rest of the string if there is anything left
    Dim l As Long: l = Len(s)
    If l < 2 Then Exit Function
    Dim c1 As Byte: c1 = Asc(Mid(s, 1, 1))
    Dim c2 As Byte: c2 = Asc(Mid(s, 2, 1))
    If l = 2 Then
        Dim ibom As Integer
        ibom = CInt("&H" & Hex2(c2) & Hex2(c1))
        If ibom = EByteOrderMark.bom_UTF_16_BE Or ibom = EByteOrderMark.bom_UTF_16_LE Then
            IsBOM = ibom: Exit Function
        End If
    End If
    Dim c3 As Byte, c4 As Byte
    Dim lbom As Long
    If l > 2 Then
        c3 = Asc(Mid(s, 3, 1))
        lbom = CLng("&H" & Hex2(c3) & Hex2(c2) & Hex2(c1))
        If Long_IsBOM(lbom) Then
            IsBOM = lbom
            rest_out = Mid(s, 4)
            Exit Function
        End If
    End If
    If l > 3 Then
        c4 = Asc(Mid(s, 4, 1))
        lbom = CLng("&H" & Hex2(c4) & Hex2(c3) & Hex2(c2) & Hex2(c1))
        If Long_IsBOM(lbom) Then
            IsBOM = lbom
            rest_out = Mid(s, 5)
            Exit Function
        End If
    End If
End Function

Public Function Long_IsBOM(ByVal Value As Long) As EByteOrderMark
    Long_IsBOM = Value
    Select Case Value
    Case EByteOrderMark.bom_UTF_8:      Exit Function
    Case EByteOrderMark.bom_UTF_16_BE:  Exit Function
    Case EByteOrderMark.bom_UTF_16_LE:  Exit Function
    Case EByteOrderMark.bom_UTF_32_BE:  Exit Function
    Case EByteOrderMark.bom_UTF_32_LE:  Exit Function
    Case EByteOrderMark.bom_UTF_7:      Exit Function
    Case EByteOrderMark.bom_UTF_1:      Exit Function
    Case EByteOrderMark.bom_UTF_EBCDIC: Exit Function
    Case EByteOrderMark.bom_SCSU:       Exit Function
    Case EByteOrderMark.bom_BOCU_1:     Exit Function
    Case EByteOrderMark.bom_GB_18030:   Exit Function
    Case Else: Long_IsBOM = bom_None
    End Select
End Function

Public Function EByteOrderMark_Parse(ByVal Value As Long) As EByteOrderMark
    
'    Dim e  As Long
'
'    e = Value
'    If e = EByteOrderMark.bom_UTF_32_BE Or _
'       e = EByteOrderMark.bom_GB_18030 Or _
'       e = EByteOrderMark.bom_UTF_EBCDIC Or _
'       e = EByteOrderMark.bom_UTF_32_LE Then EByteOrderMark_Parse = e: Exit Function
'
'    e = Value And &HFFFFFF
'    If e = EByteOrderMark.bom_SCSU Or _
'       e = EByteOrderMark.bom_UTF_8 Or _
'       e = EByteOrderMark.bom_BOCU_1 Or _
'       e = EByteOrderMark.bom_UTF_1 Then EByteOrderMark_Parse = e: Exit Function
'
'    If e = EByteOrderMark.bom_UTF_7 Then
'        e = Value \ 2 ^ 24 'shiftright 24 bits
'        If e = &H38 Or e = &H39 Or e = &H2B Or e = &H2F Then _
'                    EByteOrderMark_Parse = EByteOrderMark.bom_UTF_7: Exit Function
'    End If
'
'    e = Value And &HFFFF&
'    If e = EByteOrderMark.bom_UTF_16_BE Or _
'       e = EByteOrderMark.bom_UTF_16_LE Then EByteOrderMark_Parse = e: Exit Function
    EByteOrderMark_Parse = Long_IsBOM(Value)
End Function

Public Function EByteOrderMark_ToStr(ByVal Value As EByteOrderMark) As String
    Dim s As String
    Select Case Value
    Case EByteOrderMark.bom_GB_18030:   s = "bom_GB_18030"
    Case EByteOrderMark.bom_BOCU_1:     s = "bom_BOCU_1"
    Case EByteOrderMark.bom_SCSU:       s = "bom_SCSU"
    Case EByteOrderMark.bom_UTF_1:      s = "bom_UTF_1"
    Case EByteOrderMark.bom_UTF_16_BE:  s = "bom_UTF_16_BE"
    Case EByteOrderMark.bom_UTF_16_LE:  s = "bom_UTF_16_LE"
    Case EByteOrderMark.bom_UTF_32_BE:  s = "bom_UTF_32_BE"
    Case EByteOrderMark.bom_UTF_32_LE:  s = "bom_UTF_32_LE"
    Case EByteOrderMark.bom_UTF_7:      s = "bom_UTF_7"
    Case EByteOrderMark.bom_UTF_8:      s = "bom_UTF_8"
    Case EByteOrderMark.bom_UTF_EBCDIC: s = "bom_UTF_EBCDIC"
    End Select
    EByteOrderMark_ToStr = s
End Function

Public Function ConvertFromUTF8(ByRef Source() As Byte) As String
    'All credits for this function are going to Philipp Stephani from ActiveVB
    'http://www.activevb.de/rubriken/faq/faq0155.html
    Dim Size    As Long:       Size = UBound(Source) - LBound(Source) + 1
    Dim pSource As LongPtr: pSource = VarPtr(Source(LBound(Source)))
    Dim Length  As Long:     Length = MultiByteToWideChar(CP_UTF8, 0, pSource, Size, 0, 0)
    Dim buffer  As String:   buffer = Space$(Length)
    MultiByteToWideChar CP_UTF8, 0, pSource, Size, StrPtr(buffer), Length
    ConvertFromUTF8 = buffer
End Function

' ^ ' ############################## ' ^ '    Unicode-BOM functions    ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    Special functions    ' v ' ############################## ' v '

Public Property Get App_EXEName() As String
#If VBA6 Or VBA7 Then
    App_EXEName = Application.Name
#Else
    App_EXEName = App.EXEName
#End If
End Property

Public Function GetGreekAlphabet() As String
    'returns the greek alphabet all upper- and lower-case letters
    Dim s As String
    Dim i As Long
    Dim alp As Long: alp = 913 'the upper greek letter Alpha = 913
    For i = alp To alp + 24
        s = s & ChrW(i)
    Next
    s = s & " "
    alp = alp + 32             'the lower greek letter alpha = 945
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

Public Function GetTabbedText(s As String, Optional onlyNewLine As Boolean = False, Optional NumOnly As Boolean = False) As String
    'takes any string, first replaces any vbtab into normal space
    'then separates every value in the string with tabs, lines with vbcrlf
    Dim t As String: t = Replace(s, vbTab, " ")
    Dim lines() As String: lines = Split(t, vbCrLf)
    Dim i As Long
    'Dim onlyNewLine As Boolean: onlyNewLine = False 'Me.cbNewlineOnly.Value
    Dim svbCrLf As String: If onlyNewLine Then svbCrLf = vbCrLf
    'jeden Wert in eine neue Zeile
    For i = LBound(lines) To UBound(lines)
        Dim Line As String
        'alle mehrfachen Whitespaces enfernen
        Line = DeleteMultiWS(lines(i))
        'für Excel: alle Zahlen mit Komma(",") statt Punkt(".")
        Line = Replace(Line, ".", ",")
        If Left(Line, 3) = "K45" Then
            Debug.Assert True
        End If
        If NumOnly Then
            Dim sa() As String: sa = Split(Line, " ")
            Dim j As Long, u As Long: u = UBound(sa)
            Line = ""
            For j = 0 To u
                If IsNumeric(sa(j)) Then
                    Line = Line & sa(j) & svbCrLf
                    If onlyNewLine Then
                        'line = line & vbNewLine
                    Else
                        If j < u Then
                            Line = Line & vbTab '" "
                        End If
                    End If
                End If
            Next
        Else
            If onlyNewLine Then
                Line = Replace(Line, " ", vbCrLf)
            Else
                Line = Replace(Line, " ", vbTab)
            End If
        End If
        lines(i) = Line
    Next
    GetTabbedText = Join(lines, vbCrLf)
End Function
' ^ ' ############################## ' ^ '    Special functions    ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    Keyboard functions    ' v ' ############################## ' v '
Public Function IsAlt(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal KeyCodeConstants_KeyToCheck As Long) As Boolean
    IsAlt = (KeyCode = KeyCodeConstants_KeyToCheck) And (Shift = ShiftConstants.vbAltMask)
End Function

Public Function IsCtrl(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal KeyCodeConstants_KeyToCheck As Long) As Boolean
    IsCtrl = (KeyCode = KeyCodeConstants_KeyToCheck) And (Shift = ShiftConstants.vbCtrlMask)
End Function

Public Function IsShift(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal KeyCodeConstants_KeyToCheck As Long) As Boolean
    IsShift = (KeyCode = KeyCodeConstants_KeyToCheck) And (Shift = ShiftConstants.vbShiftMask)
End Function

Public Function IsCtrlAlt(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal KeyCodeConstants_KeyToCheck As Long) As Boolean
    IsCtrlAlt = (KeyCode = KeyCodeConstants_KeyToCheck) And (Shift = (ShiftConstants.vbCtrlMask Or ShiftConstants.vbAltMask))
End Function

Public Function IsShiftAlt(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal KeyCodeConstants_KeyToCheck As Long) As Boolean
    IsShiftAlt = (KeyCode = KeyCodeConstants_KeyToCheck) And (Shift = (ShiftConstants.vbShiftMask Or ShiftConstants.vbAltMask))
End Function

Public Function IsCtrlShift(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal KeyCodeConstants_KeyToCheck As Long) As Boolean
    IsCtrlShift = (KeyCode = KeyCodeConstants_KeyToCheck) And (Shift = (ShiftConstants.vbCtrlMask Or ShiftConstants.vbShiftMask))
End Function

Public Function IsCtrlShiftAlt(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal KeyCodeConstants_KeyToCheck As Long) As Boolean
    IsCtrlShiftAlt = (KeyCode = KeyCodeConstants_KeyToCheck) And (Shift = (ShiftConstants.vbCtrlMask Or ShiftConstants.vbShiftMask Or ShiftConstants.vbAltMask))
End Function

'Testing:
'Private Sub Form1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Select Case True
'    Case MString.IsAlt(KeyCode, Shift, vbKeyA):          MsgBox "Alt & " & ChrW(vbKeyA)
'    Case MString.IsCtrl(KeyCode, Shift, vbKeyA):         MsgBox "Ctrl & " & ChrW(vbKeyA)
'    Case MString.IsShift(KeyCode, Shift, vbKeyA):        MsgBox "Shift & " & ChrW(vbKeyA)
'    Case MString.IsCtrlAlt(KeyCode, Shift, vbKeyA):      MsgBox "Ctrl & Alt & " & ChrW(vbKeyA)
'    Case MString.IsShiftAlt(KeyCode, Shift, vbKeyA):     MsgBox "Shift & Alt & " & ChrW(vbKeyA)
'    Case MString.IsCtrlShift(KeyCode, Shift, vbKeyA):    MsgBox "Ctrl & Shift & " & ChrW(vbKeyA)
'    Case MString.IsCtrlShiftAlt(KeyCode, Shift, vbKeyA): MsgBox "Ctrl & Shift & Alt & " & ChrW(vbKeyA)
'    End Select
'End Sub

' ^ ' ############################## ' ^ '    Keyboard functions    ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    Encoding functions    ' v ' ############################## ' v '

Private Sub InitBase64() 'Code As String)
    ReDim B64(0 To 63): B64() = StrConv(Base64, vbFromUnicode)
    ' we create a second reversed table for decoding:
    ReverseCode B64, RevB64
    m_B64_IsInitialized = True
End Sub

Private Sub ReverseCode(Code() As Byte, Rev_out() As Byte)
    ReDim Rev_out(0 To 255) 'MaxValue of Code() is 255
    Dim i As Integer
    For i = 0 To UBound(Code)
        Rev_out(Code(i)) = i
    Next i
End Sub

Public Function Base64_EncodeString(Value As String) As String
    
    If Not m_B64_IsInitialized Then InitBase64
    If Len(Value) = 0 Then Exit Function
    
    Dim Source() As Byte: Source() = StrConv(Value, vbFromUnicode)
    Dim Result() As Byte
    Base64_EncodeBytes Source, Result
    
    Base64_EncodeString = StrConv(Result, vbUnicode)
    
End Function

Public Function Base64_DecodeString(Value As String) As String
    
    If Not m_B64_IsInitialized Then InitBase64
    If Len(Value) = 0 Then Exit Function
    
    Dim Source() As Byte: Source() = StrConv(Value, vbFromUnicode)
    Dim Result() As Byte
    Base64_DecodeBytes Source, Result
    
    Base64_DecodeString = StrConv(Result, vbUnicode)
    
End Function

Public Sub Base64_EncodeBytes(Source() As Byte, Result_out() As Byte)
    Dim l As Long: l = UBound(Source) - LBound(Source) + 1
    Dim rest As Long: rest = l Mod 3
    Dim n As Long
    If rest > 0 Then
        n = ((l \ 3) + 1) * 3
        ReDim Preserve Source(0 To n - 1)
    Else
        n = l
    End If
    
    ReDim Result_out(0 To n * 4 / 3 - 1) As Byte ' Das Ergebnis ist 4/3 mal so lang
    
    Dim i As Long, k As Long
    Dim c1 As Integer, c2 As Integer, c3 As Integer
    Dim W(0 To 3) As Integer
    For i = 0 To n / 3 - 1
        
        k = 3 * i 'Damit k nur einmal statt dreimal berechnet werden muss.
        c1 = Source(k + 0)   ' Je drei Byte werden gelesen
        c2 = Source(k + 1)
        c3 = Source(k + 2)
        
        W(0) = Int(c1 / 4)  ' Je 6 Bit werden extrahiert
        W(1) = (c1 And 3) * 16 + Int(c2 / 16)
        W(2) = (c2 And 15) * 4 + Int(c3 / 64)
        W(3) = (c3 And 63)
        
        k = 4 * i 'Damit k nur einmal statt viermal berechnet werden muss
        Result_out(k + 0) = B64(W(0)) ' Die 6-Bit-Werte werden nach Tabelle
        Result_out(k + 1) = B64(W(1)) ' durch Zeichen ersetzt.
        Result_out(k + 2) = B64(W(2))
        Result_out(k + 3) = B64(W(3))
        
    Next
    
    ' Je nach ueberzaehligen Bytes im Ergebnis wird dieses mit Fuellbytes aufgefuellt. Das Fuellbyte ist ein "="
    Select Case rest
    Case 0  ' OK do nothing
    Case 1: Result_out(UBound(Result_out) - 1) = 61
            Result_out(UBound(Result_out)) = 61
    Case 2: Result_out(UBound(Result_out)) = 61
    End Select
    
End Sub

Public Sub Base64_DecodeBytes(Source() As Byte, Result_out() As Byte)
    
    Dim l As Long: l = UBound(Source) - LBound(Source) + 1
    
    Dim rest As Long: rest = l Mod 4
    If rest > 0 Then ' Falls Textlaenge nicht ein Vielfaches von 4 ist
                     ' Werden einfach ein paar Nullen angehaengt.
        ReDim Preserve Source(0 To l + 4 - rest)
        l = UBound(Source) - LBound(Source) + 1
    End If
    
    ' Der String wird in ein Feld umgewandelt
    ReDim Result_out(0 To l) As Byte ' Das ist mehr Platz als benoetigt, schadet aber nicht.
    Dim w1 As Integer, w2 As Integer, w3 As Integer, w4 As Integer
    
    Dim cnt As Long
    Dim i As Long
    For i = 0 To UBound(Source) Step 4
        
        w1 = RevB64(Source(i))
        w2 = RevB64(Source(i + 1))
        w3 = RevB64(Source(i + 2))
        w4 = RevB64(Source(i + 3))
        
        Result_out(cnt) = ((w1 * 4 + Int(w2 / 16)) And 255)
        cnt = cnt + 1
        Result_out(cnt) = ((w2 * 16 + Int(w3 / 4)) And 255)
        cnt = cnt + 1
        Result_out(cnt) = ((w3 * 64 + w4) And 255)
        cnt = cnt + 1
        
    Next
    
    ReDim Preserve Result_out(cnt - 1) ' cut nulls
End Sub

'https://www.rfc-editor.org/rfc/rfc4627.txt
'2.5.  Strings
'
'   The representation of strings is similar to conventions used in the C
'   family of programming languages.  A string begins and ends with
'   quotation marks.  All Unicode characters may be placed within the
'   quotation marks except for the characters that must be escaped:
'   quotation mark, reverse solidus, and the control characters (U+0000
'   through U+001F).
'
'   Any character may be escaped.  If the character is in the Basic
'   Multilingual Plane (U+0000 through U+FFFF), then it may be
'   represented as a six-character sequence: a reverse solidus, followed
'   by the lowercase letter u, followed by four hexadecimal digits that
'   encode the character's code point.  The hexadecimal letters A though
'   F can be upper or lowercase.  So, for example, a string containing
'   only a single reverse solidus character may be represented as
'   "\u005C".
'
'   Alternatively, there are two-character sequence escape
'   representations of some popular characters.  So, for example, a
'   string containing only a single reverse solidus character may be
'   represented more compactly as "\\".
'
'   To escape an extended character that is not in the Basic Multilingual
'   Plane, the character is represented as a twelve-character sequence,
'   encoding the UTF-16 surrogate pair.  So, for example, a string
'   containing only the G clef character (U+1D11E) may be represented as
'   "\uD834\uDD1E".
'
'
'
'Crockford                    Informational                      [Page 4]
'
'RFC 4627                          JSON                         July 2006
'
'
'         string = quotation-mark *char quotation-mark
'
'         char = unescaped /
'                escape (
'                    %x22 /          ; "    quotation mark  U+0022
'                    %x5C /          ; \    reverse solidus U+005C
'                    %x2F /          ; /    solidus         U+002F
'                    %x62 /          ; b    backspace       U+0008
'                    %x66 /          ; f    form feed       U+000C
'                    %x6E /          ; n    line feed       U+000A
'                    %x72 /          ; r    carriage return U+000D
'                    %x74 /          ; t    tab             U+0009
'                    %x75 4HEXDIG )  ; uXXXX                U+XXXX
'
'         escape = %x5C              ; \
'
'         quotation-mark = %x22      ; "
'
'         unescaped = %x20-21 / %x23-5B / %x5D-10FFFF

Function JSONEscaped_Encode(ByVal Value As String) As String
    
    'escape
    ' " = \u0022
    ' \ = \u005c
    ' / = \u002f
    ' b = \u0008
    '
End Function

'JSONEscaped_Decode("")
'JSONEscaped_Decode("\u")
'JSONEscaped_Decode("\u00200")
'JSONEscaped_Decode("\u00c4\u00d6\u00dc\u00e4\u00f6\u00fc\u00df, \u00c4hren, \u00d6ltanker, \u00dcberschrift, F\u00e4rberkamille und Wilde M\u00f6hre \u00fcbernehmen die Hauptstra\u00dfe\u\u\u")
Public Function JSONEscaped_Decode(ByVal Value As String) As String
    Dim ch As String, sHex As String
    Dim cl As Long, pos As Long: pos = 1
    Dim l As Long: l = LenB(Value)
    If l = 0 Then Exit Function
    Dim pl As Long, sl As String
    Do While pos < l
        pos = InStrB(pos, Value, "\")
        If pos = 0 Then Exit Do
        ch = MidB(Value, pos + 2, 2)
        cl = AscW(ch)
        Select Case cl
        Case 92 '"\" 'it's just an escaped "\", replace " \\" with "\"
            pl = pos \ 2
            sl = Left(Value, pl)
            ch = "\"
            Value = sl & Replace(Value, "\\", ch, pl + 1, 1)
            pos = pos + 2
        Case 110 '"n" insert crlf
            pl = pos \ 2
            sl = Left(Value, pl)
            ch = vbCrLf
            Value = sl & Replace(Value, "\n", ch, pl + 1, 1)
            pos = pos + 2
        Case 116 '"t" insert tab
            pl = pos \ 2
            sl = Left(Value, pl)
            ch = vbTab
            Value = sl & Replace(Value, "\t", ch, pl + 1, 1)
            pos = pos + 2
        Case 117 '"u"
            sHex = MidB(Value, pos + 4, 8)
            If IsHex(sHex) Then
                cl = CLng("&H" & sHex)
                ch = ChrW(cl)
                'pl = Max(1, pos \ 2)
                pl = pos \ 2
                sl = Left(Value, pl)
                Value = sl & Replace(Value, "\u" & sHex, ch, pl + 1, 1)
            End If
            pos = pos + 2
        End Select
        l = LenB(Value)
    Loop
    JSONEscaped_Decode = Value
End Function

'Public Function URLEscaped_EncodeToUTF8(Value As String) As String
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'End Function

'"Spa%C3%9F" ="Spaß" in utf8
Public Function URLEscaped_DecodeFromUTF8(Value As String) As String
    'https://de.wikipedia.org/wiki/URL-Encoding
    Dim ch As String, sHex As String
    Dim cl As Long, pos As Long: pos = 1
    Dim l As Long: l = LenB(Value)
    If l = 0 Then Exit Function
    Dim pl As Long, sl As String
    Do While pos < l
        pos = InStrB(pos, Value, "%")
        If pos = 0 Then Exit Do
        ch = MidB(Value, pos, 2)
        cl = AscW(ch)
        Select Case cl
        Case &H25 '"%"
            sHex = MidB(Value, pos + 2, 4)
            If IsHex(sHex) Then
                cl = CLng("&H" & sHex)
                ch = Chr(cl)
                pl = pos \ 2
                sl = Left(Value, pl)
                Value = sl & Replace(Value, "%" & sHex, ch, pl + 1, 1)
            End If
            pos = pos + 2
        End Select
        l = LenB(Value)
    Loop
    Dim b() As Byte: b = StrConv(Value, vbFromUnicode)
    URLEscaped_DecodeFromUTF8 = MString.ConvertFromUTF8(b)
End Function

Public Function Encoding_GetString(enc As ETextEncoding, Bytes() As Byte) As String
Try: On Error GoTo Catch
    Dim n As Long: n = UBound(Bytes) - LBound(Bytes) + 1
    Dim l As Long: l = CLng(CDbl(n) * 2.2)
    Encoding_GetString = String(l, vbNullChar)
    Dim hr As Long
    Select Case enc
    Case ETextEncoding.Text_ASCIIEncoding:   hr = MultiByteToWideChar(enc, 0, VarPtr(Bytes(0)), n, StrPtr(Encoding_GetString), l)
    Case ETextEncoding.Text_UnicodeEncoding: Encoding_GetString = Bytes 'hr = MultiByteToWideChar(enc, 0, VarPtr(Bytes(0)), n, StrPtr(Encoding_GetString), L)
    Case ETextEncoding.Text_UTF32Encoding:   hr = MultiByteToWideChar(enc, 0, VarPtr(Bytes(0)), n, StrPtr(Encoding_GetString), l)
    Case ETextEncoding.Text_UTF7Encoding:    hr = MultiByteToWideChar(enc, 0, VarPtr(Bytes(0)), n, StrPtr(Encoding_GetString), l)
    Case ETextEncoding.Text_UTF8Encoding:    hr = MultiByteToWideChar(enc, 0, VarPtr(Bytes(0)), n, StrPtr(Encoding_GetString), l)
    End Select
Catch:
End Function

' ^ ' ############################## ' ^ '    Encoding functions    ' ^ ' ############################## ' ^ '

