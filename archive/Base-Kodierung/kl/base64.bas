Attribute VB_Name = "Module1"
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
Global Const Base64 = _
 "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

' Man koennte aber jede beliebige andere Anordnung der Zeichen waehlen.
Global Const codeB = _
 "1ABwCDeEoFuGyHs8ItJKL3M57NgOzPnQxRS0iTU+V9WXvYZabcdfhjkl6mp2q4r/"

Global B64() As Byte
Global Rev64() As Byte

Function base64_encode(Code() As Byte, Source As String) As String
    Dim n As Long
    Dim i As Long
    Dim c1 As Integer
    Dim c2 As Integer
    Dim c3 As Integer
    Dim w(4) As Integer
    Dim sourceB() As Byte
    Dim Result() As Byte
    Dim l As Long
    Dim k As Long
    Dim rest As Long
    Dim cnt
    
    l = Len(Source)
    If l = 0 Then
        Exit Function
    End If
    sourceB() = StrConv(Source, vbFromUnicode)
    
    rest = l Mod 3
    If rest > 0 Then
        n = ((l \ 3) + 1) * 3
        ReDim Preserve sourceB(n - 1)
    Else
        n = l
    End If
    
    ReDim Result(4 * n / 3 - 1) ' Das Ergebnis ist 4/3 mal so lang
    
    cnt = 0
    For i = 0 To n / 3 - 1
    
        k = 3 * i 'Damit k nur ein- statt dreimal berechnet werden muss.
        c1 = sourceB(k)     ' Je drei Byte werden gelesen
        c2 = sourceB(k + 1)
        c3 = sourceB(k + 2)
        
        w(1) = Int(c1 / 4)  ' Je 6 Bit werden extrahiert
        w(2) = (c1 And 3) * 16 + Int(c2 / 16)
        w(3) = (c2 And 15) * 4 + Int(c3 / 64)
        w(4) = c3 And 63
          
        k = 4 * i 'Dami k nur ein- statt viermal berechnet werden muss
        Result(k) = B64(w(1))     ' Die 6-Bit-Werte werden nach Tabelle
        Result(k + 1) = B64(w(2)) ' durch Zeichen ersetzt.
        Result(k + 2) = B64(w(3))
        Result(k + 3) = B64(w(4))
    
    Next i
    
    ' Je nach ueberzaehligen Bytes im Ergebnis wird dieses
    ' Fuellbytes aufgefuellt. Das Fuellbyte ist ein "="
    
    Select Case rest
    
    Case 0
    ' nix tun
    Case 1
        
        Result(UBound(Result)) = 61
        Result(UBound(Result) - 1) = 61
    Case 2
        '
        Result(UBound(Result)) = 61
    End Select
    
    base64_encode = StrConv(Result, vbUnicode)

End Function

Function base64_decode(Code() As Byte, Source As String) As String

    On Error GoTo err1
    
    Dim n As Long
    Dim w1 As Integer
    Dim w2 As Integer
    Dim w3 As Integer
    Dim w4 As Integer
    Dim sourceB() As Byte
    Dim Result() As Byte
    Dim l As Long
    Dim rest As Long
    Dim cnt As Long
    
    l = Len(Source)
    If l = 0 Then
        Exit Function
    End If
    
    rest = l Mod 4
    If rest > 0 Then ' Falls Textlaenge nicht ein Vielfaches von 4 ist
                     ' Werden einfach ein paar Nullen angehaengt.
        Source = Source + String$(4 - rest, 0)
        l = Len(Source)
    End If
    
    ' Der String wird in ein Feld umgewandelt
    sourceB() = StrConv(Source, vbFromUnicode)
    ReDim Result(l) ' Das ist mehr Platz als benoetigt, schadet aber nicht.
   
    For n = 0 To UBound(sourceB) Step 4
        w1 = Code(sourceB(n))
        w2 = Code(sourceB(n + 1))
        w3 = Code(sourceB(n + 2))
        w4 = Code(sourceB(n + 3))
        
        Result(cnt) = ((w1 * 4 + Int(w2 / 16)) And 255)
        cnt = cnt + 1
        Result(cnt) = ((w2 * 16 + Int(w3 / 4)) And 255)
        cnt = cnt + 1
        Result(cnt) = ((w3 * 64 + w4) And 255)
        cnt = cnt + 1
    Next n
   
done:

    ReDim Preserve Result(cnt - 1) ' Nullen abschneiden
    ' und zurueck in String verwandeln.
    base64_decode = StrConv(Result, vbUnicode)
    Exit Function
   
err1:
    Select Case Err
   
    Case 9
        ' Dies sollte eigentlich nicht passieren...
        Resume Next
        
    Case Else
        MsgBox Error
    
    End Select

End Function

Sub IniCode(Code As String)
    ReDim B64(63)
    ' Die Austauschtabelle wird in ein Bytearray uebertragen.
    B64() = StrConv(Code, vbFromUnicode)
    
    ' Und hier wird eine 2. umgekehrte Tabelle fuer die Dekodierung
    ' erstellt. Dies ist schneller, als die Tabelle
    ' jeweils nach dem Zeichen zu durchsuchen.
    Call ReverseCode(B64, Rev64)
End Sub

Function RemoveCRLF(text As String) As String
    Dim OutText As String
    Dim Oneline As String
    Dim pos1 As Long
    Dim pos2 As Long
    
    ' Carriage-Return und Line-Feed koennen per Definition
    ' nicht in einem mit Base64 kodierten Text enthalten sein.
    ' Sie werden aber meist nach je 45-60 Zeichen eingefuegt,
    ' um den Text lesbar zu machen. Hier werden sie wieder entfernt.
    
    pos1 = 1
    Do
        
        pos2 = InStr(pos1, text, vbCrLf, 1)
        If pos2 > 0 Then
            Oneline = Mid$(text, pos1, pos2 - pos1)
            OutText = OutText + Oneline
            pos1 = pos2 + 2
        Else
            Oneline = Mid$(text, pos1)
            OutText = OutText + Oneline
        End If
    
    Loop Until pos2 = 0
    
    RemoveCRLF = OutText
End Function

Function TextBlock(text As String, ByVal nChars As Long) As String
    ' Erzeugung eines Textblockes mit konstanter
    ' Zeilenlaenge fuer die Darstellung. Dies wird bei
    ' Mailattachments auch gemacht.

    Dim OutText As String
    Dim Oneline As String
    Dim i As Long
    
    For i = 1 To Len(text) Step nChars
        Oneline = Mid$(text, i, nChars) + vbCrLf
        OutText = OutText + Oneline
    Next i

    TextBlock = OutText
End Function

Sub ReverseCode(Code() As Byte, Rev() As Byte)
    Dim i As Integer
    ReDim Rev(255) '255 ist der maximale Wert der auftauchen koennte.
    
    For i = 0 To UBound(Code)
        Rev(Code(i)) = i
    Next i
    
    ' Rev() wird modifiziert zureuckgegeben, da wir es Byref
    ' uebergeben haben
End Sub
