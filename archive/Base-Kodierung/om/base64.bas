Attribute VB_Name = "MString"
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.
'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !
' Autor: K. Langbein Klaus@ActiveVB.de
' Beschreibung:
' Demonstration der Base64-Kodierung und Dekodierung
' Dies ist die richtige Austauschtabelle fuer Base64.
Option Explicit
                        '65          -         90, 97         -         122, 48 - 57, 43, 47
Private Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private m_B64_IsInitialized As Boolean
Private B64() As Byte
Private RevB64() As Byte

Private Sub InitBase64() 'Code As String)
    ReDim B64(0 To 63)
    B64() = StrConv(Base64, vbFromUnicode)
    ' we create a second reversed table for decoding
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

'Function base64_encode(Code() As Byte, Source As String) As String
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
    
    'Dim cnt As Long: cnt = 0
    Dim i As Long, k As Long
    Dim c1 As Integer, c2 As Integer, c3 As Integer
    Dim w(0 To 3) As Integer
    For i = 0 To n / 3 - 1
        
        k = 3 * i 'Damit k nur einmal statt dreimal berechnet werden muss.
        c1 = Source(k + 0)   ' Je drei Byte werden gelesen
        c2 = Source(k + 1)
        c3 = Source(k + 2)
        
        w(0) = Int(c1 / 4)  ' Je 6 Bit werden extrahiert
        w(1) = (c1 And 3) * 16 + Int(c2 / 16)
        w(2) = (c2 And 15) * 4 + Int(c3 / 64)
        w(3) = (c3 And 63)
        
        k = 4 * i 'Damit k nur einmal statt viermal berechnet werden muss
        Result_out(k + 0) = B64(w(0)) ' Die 6-Bit-Werte werden nach Tabelle
        Result_out(k + 1) = B64(w(1)) ' durch Zeichen ersetzt.
        Result_out(k + 2) = B64(w(2))
        Result_out(k + 3) = B64(w(3))
        
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

Public Function RemoveCRLF(text As String) As String
    
    Dim s As String
    Dim Oneline As String
    
    ' Carriage-Return und Line-Feed koennen per Definition
    ' nicht in einem mit Base64 kodierten Text enthalten sein.
    ' Sie werden aber meist nach je 45-60 Zeichen eingefuegt,
    ' um den Text lesbar zu machen. Hier werden sie wieder entfernt.
    
    Dim pos1 As Long: pos1 = 1
    Dim pos2 As Long
    Do
        
        pos2 = InStr(pos1, text, vbCrLf, 1)
        If pos2 > 0 Then
            Oneline = Mid$(text, pos1, pos2 - pos1)
            s = s & Oneline
            pos1 = pos2 + 2
        Else
            Oneline = Mid$(text, pos1)
            s = s + Oneline
        End If
    
    Loop Until pos2 = 0
    
    RemoveCRLF = s
End Function

Public Function TextBlock(text As String, ByVal nChars As Long) As String
    ' Erzeugung eines Textblockes mit konstanter
    ' Zeilenlaenge fuer die Darstellung. Dies wird bei
    ' Mailattachments auch gemacht.

    Dim s As String
    Dim Oneline As String
    Dim i As Long
    
    For i = 1 To Len(text) Step nChars
        Oneline = Mid$(text, i, nChars) & vbCrLf
        s = s + Oneline
    Next i
    
    TextBlock = s
End Function
