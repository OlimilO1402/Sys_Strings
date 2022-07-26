Attribute VB_Name = "Module2"
Option Explicit

'
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytLen As Long)
Private Declare Sub RtlMoveMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (pDst As String, pSrc As Any, ByVal bytLen As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (pv As Any)

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-lstrlenw
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-lstrlena
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-lstrcpyw
'Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpCWStrDst As Long, ByVal lpCWStrSrc As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-lstrcpya
'Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpCStrDst As Long, ByVal lpCStrSrc As Long) As Long

'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-lstrcpyna
'Private Declare Function lstrcpynA Lib "kernel32" (lpStrDst As String, ByVal lpStrSrc As Long, ByVal iMaxLength As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-lstrcpynw
'Private Declare Function lstrcpynW Lib "kernel32" (ByVal lpStrDst As Long, ByVal lpStrSrc As Long, ByVal iMaxLength As Long) As Long

'STRSAFEAPI StringCchCopyA(
'  [out] STRSAFE_LPSTR  pszDest,
'  [in]  size_t         cchDest,
'  [in]  STRSAFE_LPCSTR pszSrc
');
'https://docs.microsoft.com/en-us/windows/win32/api/strsafe/nf-strsafe-stringcchcopya
Private Declare Function StringCchCopyA Lib "strsafe" (ByVal pszDst As Long, ByVal cchDst As Long, ByVal pszSrc As Long) As Long

'https://docs.microsoft.com/en-us/windows/win32/api/strsafe/nf-strsafe-stringcchcopyw
Private Declare Function StringCchCopyW Lib "strsafe" (ByVal pszDst As Long, ByVal cchDst As Long, ByVal pszSrc As Long) As Long


' ----==== Pointer to String ====----
'copies from a pointer to a WIDE-string
Public Function StringFromLPCWStr(ByVal lpCWStr As Long) As String
    
    ' Länge des Strings in zeichen im Speicher
    Dim lLen As Long: lLen = lstrlenW(lpCWStr)
    
    ' gleich raus wenn Länge = 0
    If lLen = 0 Then Exit Function
    
    ' RückgabeString dimensionieren
    StringFromLPCWStr = Space(lLen)
    
    ' String vom Pointer in den Puffer kopieren
    RtlMoveMemory ByVal StrPtr(StringFromLPCWStr), ByVal lpCWStr, lLen * 2
    'Dim hr As Long
    'hr = StringCchCopyW(StrPtr(StringFromLPCWStr), lLen, lpCWStr)
    
    'oder
    'lstrcpynW StrPtr(StringFromLPCWStr), lpCWStr, lLen
    
    ' String im Speicher freigeben
    'CoTaskMemFree lpCWStr
    
End Function

'copies from an ANSI string-pointer
Public Function StringFromLPCStr(ByVal lpCStr As Long) As String
    
    ' Länge des Strings in zeichen im Speicher
    Dim lLen As Long: lLen = lstrlenA(lpCStr)
    
    ' gleich raus wenn Länge = 0
    If lLen = 0 Then Exit Function
    
    ' RückgabeString dimensionieren
    StringFromLPCStr = Space(lLen / 2)
    
    ' String vom Pointer in den Puffer kopieren
    RtlMoveMemoryStr StringFromLPCStr, ByVal lpCStr, lLen
    'StringFromLPCStr = StrConv(StringFromLPCStr, vbUnicode)
    'oder
    'lstrcpynA StringFromLPCStr, lpCStr, lLen
    
    ' String im Speicher freigeben
    'CoTaskMemFree lpCStr
    
End Function


